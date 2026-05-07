const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('menu_master.xlsx');
const sheetName = workbook.SheetNames[0];
const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

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
        wa_item: row['wa item']
    });
});

const finalMenuArray = Object.values(menuGroups);
fs.writeFileSync('menu.json', JSON.stringify(finalMenuArray, null, 2));

console.log('Fixed! JSON now carries over titles and sections correctly.');