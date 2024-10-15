const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// List of URLs to scrape
const urlsToScrape = [
    "https://www.athome.lu/srp/?tr=buy&bedrooms_min=2&bedrooms_max=3&srf_min=40&srf_max=70&sort=price_asc&q=faee1a4a&loc=L2-luxembourg&ptypes=new-property"
    // Add more URLs as needed
];

// Delay function
function delay(time) {
    return new Promise(resolve => setTimeout(resolve, time));
}

// Function to get latitude and longitude based on the address
async function getLatLong(address) {
    const url = `https://nominatim.openstreetmap.org/search?q=${encodeURIComponent(address)}&format=json&limit=1`;

    try {
        const response = await fetch(url);
        const data = await response.json();

        if (data && data.length > 0) {
            const { lat, lon } = data[0];
            console.log(`Latitude: ${lat}, Longitude: ${lon}`);
            return { lat, lon };
        } else {
            console.error("No results found for the given address.");
            return { lat: null, lon: null };
        }
    } catch (error) {
        console.error("Error fetching data:", error);
        return { lat: null, lon: null };
    }
}

// Function to sanitize filenames
const sanitizeFilename = (url) => {
    return url
        .replace("https://", "")
        .replace(/\//g, "_")
        .replace(/[<>:"/\\|?*&=]/g, "_"); // Replaces invalid characters with an underscore
};

// Function to get transaction type
async function getTransactionType(page) {
    const transactionType = await page.$eval('a.handle', el => el.textContent.trim().toLowerCase());
    return transactionType;
}

// Function to get property data
async function getPropertyData(page, propertyUrl, transactionType) {
    await page.goto(propertyUrl, { waitUntil: 'load', timeout: 0 });

    // Check if the element exists
    const nameElement = await page.$('meta[name="og:title"]');
    let name = null;
    if (nameElement) {
        name = await page.$eval('meta[name="og:title"]', el => el.content);
    } else {
        console.error(`Element 'meta[name="og:title"]' not found on page ${propertyUrl}`);
    }

    const description = await page.$eval('div.collapsed p', el => el.textContent.trim(), '').catch(() => null);

    
    const address = await page.$eval('div.block-localisation-address', el => el.textContent.trim(), '').catch(() => null);
    const price = await page.$eval('span.property-card-price', el => el.textContent.trim(), '').catch(() => null);

    const characteristics = {};
    const charItems = await page.$$eval('ul.property-card-info-icons li', items => 
        items.map(item => {
            const iconClass = item.querySelector('i').className;
            const text = item.querySelector('span').textContent.trim();
            return { iconClass, text };
        })
    );

    for (let char of charItems) {
        if (char.iconClass.includes('icon-agency_area02')) {
            characteristics['Surface'] = char.text;  // E.g., 'From 30 to 113 m²'
        } else if (char.iconClass.includes('icon-agency_bed02')) {
            characteristics['Bedrooms'] = char.text;  // E.g., '0 to 3'
        } else if (char.iconClass.includes('icon-agency_room')) {
            characteristics['Rooms'] = char.text;  // E.g., '0 to 2'
        }
    }


    // Get latitude and longitude based on the address
    const { lat: latitude, lon: longitude } = await getLatLong(address);

    // Determine property type based on keywords
    const propertyTypeKeywords = [
        "maison", "appartement", "chambre", "studio", "penthouse", "duplex", "triplex",
        "loft", "mansarde", "rez-de-chaussée", "projet neuf", "résidence", "lotissement",
        "terrain", "garage", "bureau", "commerce", "local", "restaurant", "hôtel",
        "entrepôt", "exploitation agricole"
    ];
    const propertyType = propertyTypeKeywords.find(word => name && name.toLowerCase().includes(word)) || null;

    return {
        'URL' : propertyUrl,
        "Name": name,
        "Description": description, 
        "Address": address,
        "Price": price,
        "Area": characteristics["Surface"] || "",
        "Characteristics": characteristics,
        "Property Type": propertyType,
        "Transaction Type": transactionType,
        "Latitude": latitude,
        "Longitude": longitude
    };
}

// Function to scrape a page
async function scrapePage(page, url) {
    await page.goto(url, { waitUntil: 'load', timeout: 0 });

    const propertyUrls = await page.$$eval('a.property-card-link.property-price', links => links.map(link => link.href));
    return propertyUrls;
}

// Function to save Excel file dynamically in the same folder as the script
async function saveToExcel(data, startUrl) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Properties');

    worksheet.columns = [
        { header: 'URL', key: 'URL', width: 30 },
        { header: 'Name', key: 'Name', width: 30 },
        { header: 'Description', key: 'Description', width: 50 },
        { header: 'Address', key: 'Address', width: 30 },
        { header: 'Price', key: 'Price', width: 15 },
        { header: 'Area', key: 'Area', width: 15 },
        { header: 'Characteristics', key: 'Characteristics', width: 50 },
        { header: 'Property Type', key: 'Property Type', width: 20 },
        { header: 'Transaction Type', key: 'Transaction Type', width: 20 },
        { header: 'Latitude', key: 'Latitude', width: 15 },
        { header: 'Longitude', key: 'Longitude', width: 15 }
    ];

    // Add data to worksheet
    data.forEach(row => worksheet.addRow(row));

    // Dynamically generate the file name
    const sanitizedFilename = sanitizeFilename(startUrl) + ".xlsx";

    // Create output directory if it doesn't exist
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    // Save the file in the output directory
    const filePath = path.join(outputDir, sanitizedFilename);

    // Write the Excel file
    await workbook.xlsx.writeFile(filePath);

    console.log(`Data saved to ${filePath}`);
}
// Function to scrape all data for a URL
async function scrapeAllDataForUrl(browser, startUrl) {
    const page = await browser.newPage();

    // Navigate to the start URL
    await page.goto(startUrl, { waitUntil: 'load', timeout: 0 });

    try {
        // Wait for the element to appear (add a timeout of 10 seconds)
        await page.waitForSelector('header.block-alert h2', { timeout: 10000 });
        const numberOfListingsText = await page.$eval('header.block-alert h2', el => el.textContent);
        const numberOfListings = parseInt(numberOfListingsText.match(/\d[\d,.]*/)[0].replace(',', ''), 10);

        // Calculate the total number of pages
        const totalPages = Math.ceil(numberOfListings / 20);

        // Get transaction type
        const transactionType = await getTransactionType(page);

        const allPropertyUrls = [];
        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            const pageUrl = `${startUrl}&page=${pageNum}`;
            console.log(`Scraping page: ${pageUrl}`);
            const propertyUrls = await scrapePage(page, pageUrl);
            allPropertyUrls.push(...propertyUrls);
        }

        const allPropertyData = [];
        for (let propertyUrl of allPropertyUrls) {
            console.log(`Scraping property: ${propertyUrl}`);

            // Add delay of 3 seconds before scraping each property
            await delay(3000);

            const propertyData = await getPropertyData(page, propertyUrl, transactionType);
            console.log(propertyData);
            allPropertyData.push(propertyData);
        }

        // Save to Excel
        await saveToExcel(allPropertyData, startUrl);

    } catch (error) {
        console.error(`Error occurred while scraping URL: ${startUrl}. Details: ${error.message}`);
    } finally {
        await page.close();
    }
}


// Main function to scrape all URLs
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

    for (let url of urlsToScrape) {
        await scrapeAllDataForUrl(browser, url);
    }

    await browser.close();
})();
