const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

(async () => {
   const browser = await puppeteer.launch({
        headless: 'new', 
        args: [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36",
        ],
        defaultViewport: null,
    });

    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 });

    const baseUrl = 'https://www.kv.ee/search?orderby=ob&deal_type=20';
    await page.goto(baseUrl);

    const propertyType = await page.evaluate(() => {
        const activeElement = document.querySelector('div.main-menu a.active');
        return activeElement ? activeElement.querySelector('span').innerText.trim() : null;
    });

    await page.waitForSelector('span.large.stronger');

    const totalListings = await page.evaluate(() => {
        const listingElement = document.querySelector('span.large.stronger');
        const listingsText = listingElement ? listingElement.innerText : '';
        const totalListings = listingsText.match(/\d[\d\s]*\d/)[0].replace(/\s/g, '');
        return totalListings;
    });

    const listingsPerPage = 50;
    const totalPages = Math.ceil(totalListings / listingsPerPage);
    console.log('Total Pages:', totalPages);

    let allPropertyUrls = [];

    for (let pageNum = 0; pageNum < totalPages; pageNum++) {
        const currentPageUrl = `${baseUrl}&start=${pageNum * listingsPerPage}`;
        await page.goto(currentPageUrl, { waitUntil: 'networkidle2' });
        await page.waitForSelector('div.description a[data-key]');
        
        const propertyUrls = await page.evaluate(() => {
            const descriptionDivs = document.querySelectorAll('div.description');
            const links = [];
            descriptionDivs.forEach(div => {
                const aTags = div.querySelectorAll('a[data-key][href]');
                aTags.forEach(aTag => {
                    const href = aTag.href;
                    if (!href.includes('/object/images')) {
                        links.push(href);
                    }
                });
            });
            return links;
        });

        allPropertyUrls = allPropertyUrls.concat(propertyUrls);
    }

    console.log('Total Property URLs:', allPropertyUrls.length);

    const propertyData = [];

    for (const propertyUrl of allPropertyUrls) {
        let retries = 3;
        let success = false;

        while (retries > 0 && !success) {
            try {
                await page.goto(propertyUrl, { waitUntil: 'networkidle2', timeout: 300000 });
                await page.waitForSelector('meta[property="og:title"]', { timeout: 100000 });

                const data = await page.evaluate(() => {
                    const getMetaContent = (property) => {
                        const metaTag = document.querySelector(`meta[property="${property}"]`);
                        return metaTag ? metaTag.content : null;
                    };

                    const extractCoordinates = () => {
                        const mapLink = document.querySelector('a[title="Suurem kaart"]');
                        if (mapLink) {
                            const href = mapLink.href;
                            const coordsMatch = href.match(/query=([\d.]+),([\d.]+)/);
                            if (coordsMatch) {
                                return {
                                    latitude: parseFloat(coordsMatch[1]),
                                    longitude: parseFloat(coordsMatch[2])
                                };
                            }
                        }
                        return { latitude: null, longitude: null };
                    };

                    const name = getMetaContent('og:title');
                    if (name) {
                        const nameParts = name.split(' - ');
                        address = nameParts.length > 1 ? nameParts[1] : 'Address not available';
                    }


                    const description = getMetaContent('og:description');
                    let price = '';

                    const priceOuterDiv = document.querySelector('div.price-outer');
                    if (priceOuterDiv) {
                        const newPriceElement = priceOuterDiv.querySelector('div.red + div');
                        if (newPriceElement) {
                            price = newPriceElement.innerText.split(' ')[0];
                        } else {
                            const priceElement = priceOuterDiv.querySelector('div');
                            price = priceElement ? priceElement.innerText.split(' ')[0] : '';
                        }
                    }

                    const rangePriceElement = document.querySelector('h4.strong');
                    if (rangePriceElement) {
                        const priceText = rangePriceElement.innerText.trim().replace(/\s/g, '');
                        const priceMatch = priceText.match(/(\d+[\d\s]*\d+)€/g);
                        if (priceMatch && priceMatch.length === 2) {
                            price = `${priceMatch[0].replace('€', '')} - ${priceMatch[1].replace('€', '')}`;
                        } else if (priceMatch && priceMatch.length === 1) {
                            price = priceMatch[0].replace('€', '');
                        }
                    }

                    const characteristics = {};
                    const rows = document.querySelectorAll('table.table-lined tr');

                    rows.forEach(row => {
                        const th = row.querySelector('th');
                        const td = row.querySelector('td');
                        if (th && td) {
                            const key = th.innerText.trim();
                            const value = td.innerText.trim();
                            characteristics[key] = value;
                        }
                    });

                    const area = characteristics['Üldpind'] ? characteristics['Üldpind'].replace('m²', '').trim() : null;
                    const coordinates = extractCoordinates();

                    let transactionType = '';
                    if (price && parseFloat(price.replace(/[^\d]/g, '')) > 10000) {
                        transactionType = 'sale';
                    } else {
                        transactionType = 'rent';
                    }

                    return {
                        propertyUrl: window.location.href,
                        name,
                        address,
                        price,
                        description,
                        area,
                        characteristics: JSON.stringify(characteristics),
                        latitude: coordinates.latitude,
                        longitude: coordinates.longitude,
                        transactionType
                    };
                });

                data.propertyType = propertyType;

                propertyData.push(data);
                console.log(`Data for URL: ${propertyUrl}`, data);

                success = true;

            } catch (error) {
                console.error(`Error fetching data for URL: ${propertyUrl}. Retries left: ${retries - 1}`);
                retries--;
                if (retries === 0) {
                    console.error(`Failed to fetch data for URL: ${propertyUrl} after 3 attempts.`);
                }
            }
        }
    }

    await browser.close();

    const urlParts = new URL(baseUrl);
    const cleanUrl = `${urlParts.host.replace(/\./g, '_')}_${urlParts.pathname.replace(/\//g, '_')}` + 
                    (urlParts.search ? `_${urlParts.search.replace(/[\/\\?%*:|"<>]/g, '_')}` : '');
    const fileName = `${cleanUrl}.xlsx`;

    // Ensure the 'output' directory exists before saving the file
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    const outputFilePath = path.join(outputDir, fileName);

    // Create an Excel workbook and add a sheet with the specified columns
    const ws = xlsx.utils.json_to_sheet(propertyData, {
        header: ['propertyUrl', 'name', 'address', 'price', 'description', 'area', 'characteristics', 'longitude', 'latitude', 'transactionType', 'propertyType']
    });
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, ws, 'Properties');
    xlsx.writeFile(wb, outputFilePath);

    console.log(`Data saved to ${outputFilePath}`);

})();
