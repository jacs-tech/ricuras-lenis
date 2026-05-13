const XLSX = require('xlsx');
const fs = require('fs');

// --- PROJECT SETTINGS ---
// Just change these three lines for every new store project
const EXCEL_FILE = 'menu_lenis.xlsx'; 
const MENU_SHEET = 'Menu';
const CONFIG_SHEET = 'Config';

// --- EXECUTION ---
if (!fs.existsSync(EXCEL_FILE)) {
    console.error(`❌ Error: El archivo ${EXCEL_FILE} no existe.`);
    process.exit(1);
}

const workbook = XLSX.readFile(EXCEL_FILE);

// --- NEW PART: STORE INFO & SCHEDULE ---
// Assuming Sheet 2 is named 'Config' (or you can use workbook.SheetNames[1])
const configSheet = workbook.Sheets[CONFIG_SHEET];
const configRaw = XLSX.utils.sheet_to_json(configSheet, { header: 1 });

const config = {
    whatsapp: configRaw[1][1],
    map_url: configRaw[2][1],
    // NEW: Capture cell B15 (Row 15, Index 14)
    chargePackage: (configRaw[14] && configRaw[14][1] === "yes"), 
    // New Automation Fields
    storeName: configRaw[17] ? configRaw[17][1] : 'Mi Tienda',    // Cell B18
    repoName: configRaw[18] ? configRaw[18][1] : '',
    horario: []   
};

// Loop through the schedule rows (Index 5 to 11)
for (let i = 5; i <= 11; i++) {
    if (configRaw[i]) { 
        let openVal = configRaw[i][2];
        let closeVal = configRaw[i][3];
        let promoVal = configRaw[i][4]; // <--- NEW: Column E (Index 4)

        config.horario.push({
            open: (openVal === undefined || openVal === "null") ? null : openVal,
            close: (closeVal === undefined || closeVal === "null") ? null : closeVal,
            isPromo: promoVal === "yes" // <--- NEW: Stores true if "yes", false otherwise
        });
    }
}

fs.writeFileSync('config.json', JSON.stringify(config, null, 2));
console.log('✅ config.json generated successfully!');
// --- END OF NEW PART ---

// YOUR EXISTING MENU CODE STARTS HERE
const menuSheet = workbook.Sheets[MENU_SHEET];
if (!menuSheet) {
    console.error("Error: No se encontró la hoja 'Menu'");
    process.exit(1);
}

const rawData = XLSX.utils.sheet_to_json(menuSheet);
const menuGroups = {};
let lastSection = "";
let lastTitle = "";

rawData.forEach(row => {
    // If the cell is empty, use the one from the row above
    const currentSection = row.section || lastSection;
    const currentTitle = row.title || lastTitle;

    // Update our "memory" for the next row
    lastSection = currentSection;
    lastTitle = currentTitle;

    if (!menuGroups[currentSection]) {
        menuGroups[currentSection] = {
            section: currentSection,
            title: currentTitle,
            items: []
        };
    }

    menuGroups[currentSection].items.push({
        nickname: row.nickname,
        name: row.item,
        desc: row.desc,
        price: row.price,
        wa_item: row['wa item'],
        // NEW: Capture Column H (Package price)
        packagePrice: row.package || 0 
    });
});

const finalMenuArray = Object.values(menuGroups);
fs.writeFileSync('menu.json', JSON.stringify(finalMenuArray, null, 2));

console.log('Fixed! JSON now carries over titles and sections correctly.');