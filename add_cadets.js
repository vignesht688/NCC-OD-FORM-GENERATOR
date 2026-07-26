const fs = require('fs');
const path = require('path');

// Absolute path to the cadets file in the workspace
const CADETS_FILE = 'c:\\Users\\jayav\\OneDrive\\Desktop\\od form for ncc\\cadets.json';

const newCadetsList = [
  { "register_no": "", "name": "ABI DHARSHAN H K", "department": "BSC MATHS", "shift": "" },
  { "register_no": "", "name": "VENKATESH P", "department": "BSC MATHS", "shift": "" },
  { "register_no": "", "name": "KRISHANT R", "department": "BSC MATHS", "shift": "" },
  { "register_no": "", "name": "GOWTHAM S", "department": "BSC CHEMISTRY", "shift": "" },
  { "register_no": "", "name": "SANTHOSH G", "department": "BSC MATHS", "shift": "" },
  { "register_no": "", "name": "EISHANTH K N", "department": "BSC PHYSICS", "shift": "" },
  { "register_no": "", "name": "HARISH B", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "DEEPAK R", "department": "BSC ECS", "shift": "" },
  { "register_no": "", "name": "KISHORE J", "department": "BCOM", "shift": "" },
  { "register_no": "", "name": "DHARSHAN P", "department": "BCOM", "shift": "" },
  { "register_no": "", "name": "SUDHARSHAN K", "department": "BCOM", "shift": "" },
  { "register_no": "", "name": "KARTHI A", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "IMMANUVEL", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "YAGARAJ", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "KAVIRAJ", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "Y. SIVANESHWARAN", "department": "BSC CS WITH COGNITIVE SYSTEMS", "shift": "" },
  { "register_no": "", "name": "SHAHIL KUMAR T", "department": "BSC ECS", "shift": "" },
  { "register_no": "", "name": "AAKASH I", "department": "BBA", "shift": "" },
  { "register_no": "", "name": "SUJAN S", "department": "BSC DCFS", "shift": "" },
  { "register_no": "", "name": "THIRUKUMARAN M", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "ABINANDH V", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "VISHNUVARTHAN", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "RISHIKUMAR", "department": "BSC CS AI DS", "shift": "" },
  { "register_no": "", "name": "SRI RATHIN", "department": "BCOM", "shift": "" },
  { "register_no": "", "name": "ABISHEK SHARMA", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "DHARUKKANNA", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "SIBANIKESH", "department": "BCA", "shift": "" },
  { "register_no": "", "name": "YASWANTH KUMAR", "department": "BCA", "shift": "" },
  { "register_no": "", "name": "PUSHPAKARTHIK", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "AKILESH S", "department": "BCA", "shift": "" },
  { "register_no": "", "name": "ROHITH K", "department": "BCA", "shift": "" },
  { "register_no": "", "name": "RISHAB A", "department": "BBA CA", "shift": "" },
  { "register_no": "", "name": "GOKUL R", "department": "BSC CT", "shift": "" },
  { "register_no": "", "name": "JAI VISHNU VARDHAN", "department": "BSC CT", "shift": "" },
  { "register_no": "", "name": "HARISH S S", "department": "BSC CT", "shift": "" },
  { "register_no": "", "name": "SUJITH D", "department": "BSC CT", "shift": "" },
  { "register_no": "", "name": "SARVANA", "department": "BCOM CS", "shift": "" },
  { "register_no": "", "name": "AADITH S", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "ATHINYA KRISHNAN", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "MATHAN M", "department": "BCOM PA", "shift": "" },
  { "register_no": "", "name": "PRAKASH L", "department": "BBA", "shift": "" },
  { "register_no": "", "name": "KARTHIKEYAN M", "department": "BCOM", "shift": "SHIFT-II" },
  { "register_no": "", "name": "SHYAM A", "department": "BCOM", "shift": "SHIFT-II" },
  { "register_no": "", "name": "ANDREW JONSLY", "department": "BCOM IB", "shift": "SHIFT-I" },
  { "register_no": "", "name": "HARIHARASUDHAN G", "department": "BCOM IB", "shift": "SHIFT-I" },
  { "register_no": "", "name": "PARUKESH", "department": "BCOM IB", "shift": "SHIFT-I" },
  { "register_no": "", "name": "VARUN C", "department": "BCOM IB", "shift": "SHIFT-I" },
  { "register_no": "", "name": "SACHIN", "department": "BCOM PA", "shift": "SHIFT-II" },
  { "register_no": "", "name": "ARAVIND M", "department": "BCOM CS", "shift": "SHIFT-I" },
  { "register_no": "", "name": "ARWIN S", "department": "BCOM CS", "shift": "SHIFT-II" },
  { "register_no": "", "name": "AJAY P", "department": "BBA LOG", "shift": "SHIFT-II" },
  { "register_no": "", "name": "RAGHUL A", "department": "BBA LOG", "shift": "SHIFT-I" },
  { "register_no": "", "name": "PRASANNA S", "department": "BCOM", "shift": "SHIFT-II" },
  { "register_no": "", "name": "DHARUN PANDI G", "department": "BCOM", "shift": "SHIFT-II" },
  { "register_no": "", "name": "LUBIN ANDERSON", "department": "BSC SPORTS", "shift": "" },
  { "register_no": "", "name": "SARAN NITHISH", "department": "BCOM CA", "shift": "" },
  { "register_no": "", "name": "JASWANTH J", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "MOHAMMED IRSHAD", "department": "BSC CHEMISTRY", "shift": "" },
  { "register_no": "", "name": "UDHAYANITHI", "department": "BSC DS", "shift": "" },
  { "register_no": "", "name": "PRAGADHISHWARAN", "department": "BCOM IB", "shift": "" },
  { "register_no": "", "name": "SIVACHIDAMBARAM", "department": "BSC CSDA", "shift": "" },
  { "register_no": "", "name": "MATHAN", "department": "BCOM IT", "shift": "" },
  { "register_no": "", "name": "SARAN S T", "department": "BCOM IT", "shift": "" },
  { "register_no": "", "name": "SARVESH S", "department": "BCOM IT", "shift": "" },
  { "register_no": "", "name": "MOHAMMED JAMAL", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "SANTHOSH V", "department": "BCOM B&I", "shift": "" },
  { "register_no": "", "name": "KAVIN SAIRAJ", "department": "BSC IT", "shift": "" },
  { "register_no": "", "name": "DHUVARAKESH P", "department": "BSC IT", "shift": "" },
  { "register_no": "", "name": "ABHIMANYU", "department": "BSC CS", "shift": "" },
  { "register_no": "", "name": "VISHNUVETRICHELVAN", "department": "BSC PHYSICS", "shift": "" },
  { "register_no": "", "name": "MOHAMMED YASIN", "department": "BCOM", "shift": "" }
];

try {
  // 1. Read existing database
  console.log('Loading existing cadets database from ' + CADETS_FILE + '...');
  const fileData = fs.readFileSync(CADETS_FILE, 'utf8');
  let cadets = JSON.parse(fileData || '[]');
  console.log(`Currently has ${cadets.length} cadets.`);

  let addedCount = 0;
  let updatedCount = 0;

  // 2. Process new cadets
  newCadetsList.forEach(newCadet => {
    // Normalise name (trim and uppercase)
    const cleanName = newCadet.name.trim().toUpperCase();

    // Check if there is an exact or near match
    // Special check for THARANI -> THARANI S update
    if (cleanName === 'THARANI S') {
      const idx = cadets.findIndex(c => c.name.trim().toUpperCase() === 'THARANI');
      if (idx !== -1) {
        cadets[idx].name = 'THARANI S';
        cadets[idx].department = 'BCOM CA';
        updatedCount++;
        console.log(`Updated entry "THARANI" to "THARANI S"`);
        return;
      }
    }

    // Generic duplicate check: matches name exactly
    const exists = cadets.some(c => c.name.trim().toUpperCase() === cleanName);
    if (!exists) {
      cadets.push({
        register_no: newCadet.register_no,
        name: cleanName,
        department: newCadet.department.trim().toUpperCase(),
        shift: newCadet.shift.trim().toUpperCase()
      });
      addedCount++;
    }
  });

  // 3. Sort cadets alphabetically by name
  cadets.sort((a, b) => a.name.localeCompare(b.name));

  // 4. Save back to file
  fs.writeFileSync(CADETS_FILE, JSON.stringify(cadets, null, 2), 'utf8');
  console.log(`Successfully completed!`);
  console.log(`Added ${addedCount} new cadets.`);
  console.log(`Updated ${updatedCount} existing cadets.`);
  console.log(`New total in cadets.json: ${cadets.length} cadets.`);

} catch (err) {
  console.error('Error occurred:', err);
}
