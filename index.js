const fs = require("fs")
const path = require("path")
const csv = require('csvtojson')
const XLSX = require("xlsx");



var emailRegex = /^[-!#$%&'*+\/0-9=?A-Z^_a-z{|}~](\.?[-!#$%&'*+\/0-9=?A-Z^_a-z`{|}~])*@[a-zA-Z0-9](-*\.?[a-zA-Z0-9])*\.[a-zA-Z](-?[a-zA-Z0-9])+$/;

function isEmailValid(email) {
    if (!email)
        return false;

    if(email.length>254)
        return false;

    var valid = emailRegex.test(email);
    if(!valid)
        return false;

    // Further checking of some things regex can't handle
    var parts = email.split("@");
    if(parts[0].length>64)
        return false;

    var domainParts = parts[1].split(".");
    if(domainParts.some(function(part) { return part.length>63; }))
        return false;

    return true;
}

async function main() {
  console.log("Starting up cleaner")
  console.log("-------------------------------------------------")


  // Read all the files we are going to clean against
  const removeFiles = fs.readdirSync("./remove")
  const emailsToRemove = []

  for (const filename of removeFiles) {
    // only work with accepted fileextentions
    if (![".csv"].includes(path.extname(filename))) continue

    let emailsInDocument = 0
    console.log("Fetching emails to remove: ", filename)
  
    if (path.extname(filename) == ".csv") {
      const jsonArray = await csv().fromFile("./remove/" + filename) 
      console.log("rows found in "+ filename  + ": ", jsonArray.length)

      // loop through all the rows
      for (const row of jsonArray) {
        for (const key in row) {
          if (isEmailValid(row[key])) {
            emailsToRemove.push(row[key].toLowerCase())
            ++emailsInDocument
          }
        }
      }
    }

    console.log("Emails collected from document: ", emailsInDocument)
    console.log("-------------------------------------------------")
  }

  console.log("emailsToRemove: ", emailsToRemove.length)

  // Read files from the source folder
  const sourceFiles = fs.readdirSync("./source")

  for (const filename of sourceFiles) {
    // only work with accepted fileextentions
    if (![".csv"].includes(path.extname(filename))) continue

    console.log("-------------------------------------------------")
    console.log("Working on file: ", filename)

    const rows = []
    let duplicatesFound = 0

    if (path.extname(filename) === ".csv") {
      const jsonArray = await csv().fromFile("./source/" + filename) 
      console.log("rows found in "+ filename  + ": ", jsonArray.length)
      const hr = [] // header row

      // Write header columns
      for (const key in jsonArray[0]) {
        hr.push(key)
      }
      rows.push(hr)
      
      for (const row of jsonArray) {
        const r = [] // row with only data
        // find the email in the row
        let email

        for (const key in row) {
          if (isEmailValid(row[key])) {
            email = row[key].toLowerCase()
          }
        }

        // check if it should be removed or saved
        if (emailsToRemove.includes(email)) {
          ++duplicatesFound
          continue
        }

        // Should be stored, let's put the row into rows
        for (const key in row) {
          r.push(row[key])
        }
        rows.push(r)
      }
    }

    // stats
    console.log("We have " + rows.length + " rows left that is going to be stored")
    console.log("We removed " + duplicatesFound + "rows from the document")

    // write to document 
    console.log("Writing to output file: " + path.parse(filename).name)
    var ws = XLSX.utils.aoa_to_sheet(rows)
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet");

    XLSX.writeFile(wb, "./out/" + path.parse(filename).name + ".xlsb")
  }
}


main()
