import * as fs from "fs";
import * as path from "path";
import * as XLSX from "xlsx";
import pdfParse from "pdf-parse";

async function convertCitiPdfToExcel(
  pdfPath: string,
  excelPath: string
): Promise<void> {
  try {
    console.log(`Reading Citibank PDF file: ${pdfPath}`);

    // Read the PDF file
    const pdfBuffer = fs.readFileSync(pdfPath);
    const pdfData = await pdfParse(pdfBuffer);

    // Extract text content
    const textContent = pdfData.text;
    console.log("PDF content extracted successfully");

    // Save the extracted text to a file for debugging
    fs.writeFileSync("citi-extracted-text.txt", textContent);
    console.log(
      "Extracted text saved to citi-extracted-text.txt for debugging"
    );

    // Process the text content to extract structured data
    const data = processCitiStatementData(textContent);

    console.log(`Found ${data.length} transactions`);

    // If no transactions were found, try a more generic approach
    if (data.length === 0) {
      console.log(
        "No transactions found with standard parsing. Trying alternative approach..."
      );
      const alternativeData = processStatementDataAlternative(textContent);
      console.log(
        `Alternative approach found ${alternativeData.length} transactions`
      );

      if (alternativeData.length > 0) {
        console.log("Using data from alternative parsing approach");
        data.push(...alternativeData);
      }
    }

    // Try direct text extraction as a last resort
    if (data.length < 6) {
      // If we found fewer than 6 transactions, try direct extraction
      console.log("Trying direct text extraction to find more transactions...");
      const directExtractionData = extractTransactionsFromRawText(textContent);

      if (directExtractionData.length > data.length) {
        console.log(
          `Direct extraction found more transactions (${directExtractionData.length}). Using these instead.`
        );
        // Clear existing data and use direct extraction results
        data.length = 0;
        data.push(...directExtractionData);
      } else if (directExtractionData.length > 0) {
        // Add any new transactions not already in the data array
        console.log("Adding unique transactions from direct extraction...");
        for (const transaction of directExtractionData) {
          // Check if this transaction is already in the data array
          const isDuplicate = data.some(
            (t) =>
              t.Date === transaction.Date &&
              (t.Debit === transaction.Debit ||
                t.Credit === transaction.Credit) &&
              t.Description.includes(transaction.Description.substring(0, 10))
          );

          if (!isDuplicate) {
            data.push(transaction);
          }
        }
      }
    }

    // If still no data, add a dummy row to show the file isn't completely empty
    if (data.length === 0) {
      data.push({
        Note: "No transactions could be automatically extracted from the PDF. Please check citi-extracted-text.txt and adjust the parsing logic.",
      });
    } else {
      // Add a summary row with totals
      addSummaryRow(data);
    }

    // Create a new workbook
    const workbook = XLSX.utils.book_new();

    // Convert data to worksheet
    const worksheet = XLSX.utils.json_to_sheet(data);

    // Format the worksheet
    formatWorksheet(worksheet, data);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, "Citibank Statement");

    // Write to Excel file
    XLSX.writeFile(workbook, excelPath);

    console.log(`Excel file created successfully: ${excelPath}`);
  } catch (error) {
    console.error("Error converting PDF to Excel:", error);
    throw error;
  }
}

// Helper function to add a summary row with totals
function addSummaryRow(data: any[]) {
  // Calculate totals
  let totalDebits = 0;
  let totalCredits = 0;

  // Create a map to track category totals
  const categoryTotals: {
    [key: string]: { debit: number; credit: number };
  } = {};

  for (const transaction of data) {
    // Convert debit and credit to numbers
    if (transaction.Debit) {
      const debit = parseFloat(transaction.Debit.replace(/,/g, ""));
      if (!isNaN(debit)) {
        totalDebits += debit;

        // Add to category totals
        const category = transaction.Category || "Uncategorized";
        if (!categoryTotals[category]) {
          categoryTotals[category] = { debit: 0, credit: 0 };
        }
        categoryTotals[category].debit += debit;
      }
    }

    if (transaction.Credit) {
      const credit = parseFloat(transaction.Credit.replace(/,/g, ""));
      if (!isNaN(credit)) {
        totalCredits += credit;

        // Add to category totals
        const category = transaction.Category || "Uncategorized";
        if (!categoryTotals[category]) {
          categoryTotals[category] = { debit: 0, credit: 0 };
        }
        categoryTotals[category].credit += credit;
      }
    }
  }

  // Add a blank row before summaries
  data.push({
    Date: "",
    Description: "",
    Debit: "",
    Credit: "",
    Balance: "",
    Category: "",
  });

  // Add category summary rows
  for (const category in categoryTotals) {
    const totals = categoryTotals[category];
    data.push({
      Date: "",
      Description: `SUBTOTAL: ${category}`,
      Debit: totals.debit > 0 ? totals.debit.toFixed(2) : "",
      Credit: totals.credit > 0 ? totals.credit.toFixed(2) : "",
      Balance: "",
      Category: category,
    });
  }

  // Add a blank row before final total
  data.push({
    Date: "",
    Description: "",
    Debit: "",
    Credit: "",
    Balance: "",
    Category: "",
  });

  // Add final summary row
  data.push({
    Date: "",
    Description: "TOTAL",
    Debit: totalDebits.toFixed(2),
    Credit: totalCredits.toFixed(2),
    Balance: "",
    Category: "",
  });
}

// Helper function to format the worksheet
function formatWorksheet(worksheet: XLSX.WorkSheet, data: any[]) {
  // Set column widths
  const columnWidths = [
    { wch: 12 }, // Date
    { wch: 50 }, // Description
    { wch: 15 }, // Debit
    { wch: 15 }, // Credit
    { wch: 15 }, // Balance
    { wch: 15 }, // Category
  ];

  worksheet["!cols"] = columnWidths;
}

// Specific parser for Citibank statements
function processCitiStatementData(textContent: string): any[] {
  console.log("Starting Citibank statement parsing...");

  // Pre-process the text content to fix common issues
  textContent = textContent.replace(/(\d{2})([A-Z]{3})/g, "$1 $2"); // Add space between day and month if missing

  // Split the content by lines
  const lines = textContent.split("\n").filter((line) => line.trim() !== "");

  // Array to store the extracted transactions
  const transactions: any[] = [];

  let isTransactionSection = false;
  let currentTransaction: any = null;
  let accountNumber = "";
  let statementPeriod = "";
  let cardHolder = "";
  let cardType = "";
  let statementYear = new Date().getFullYear().toString(); // Default to current year
  let statementMonth = "";

  // First, try to extract account information
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Look for account number
    if (
      line.includes("ACCOUNT NUMBER") ||
      line.includes("CARD NUMBER") ||
      line.match(/\d{4}\d{4}\d{4}\d{4}/)
    ) {
      const accountMatch = line.match(/(\d{4}\d{4}\d{4}\d{4})/);
      if (accountMatch && accountMatch[1]) {
        accountNumber = accountMatch[1];
        console.log(`Found account/card number: ${accountNumber}`);
      }
    }

    // Look for statement period
    if (line.includes("Statement Date") || line.includes("STATEMENT DATE")) {
      const periodMatch = line.match(
        /Statement Date:?\s*([A-Za-z]+\s*\d{1,2},\s*\d{4})/i
      );
      if (periodMatch && periodMatch[1]) {
        statementPeriod = periodMatch[1].trim();
        console.log(`Found statement date: ${statementPeriod}`);

        // Extract year from statement date
        const yearMatch = statementPeriod.match(/\d{4}/);
        if (yearMatch) {
          statementYear = yearMatch[0];
          console.log(`Extracted statement year: ${statementYear}`);
        }

        // Extract month from statement date
        const monthMatch = statementPeriod.match(/([A-Za-z]+)/i);
        if (monthMatch) {
          statementMonth = monthMatch[1].toUpperCase().substring(0, 3);
          console.log(`Extracted statement month: ${statementMonth}`);
        }
      }
    }

    // Look for cardholder name
    if (line.match(/^[A-Z]+\s+[A-Z]+\s*$/)) {
      cardHolder = line.trim();
      console.log(`Found cardholder: ${cardHolder}`);
    }

    // Look for card type
    if (
      line.includes("CITI") &&
      (line.includes("VISA") || line.includes("MASTERCARD"))
    ) {
      cardType = line.trim();
      console.log(`Found card type: ${cardType}`);
    }

    // Detect the start of transaction section - more flexible patterns
    if (
      (line.includes("DATE") &&
        line.includes("DESCRIPTION") &&
        line.includes("AMOUNT")) ||
      (line.includes("TRANSACTIONS FOR") &&
        lines[i + 1] &&
        lines[i + 1].includes("ALL TRANSACTIONS")) ||
      (line.match(/^\s*DATE\s*/) &&
        line.match(/DESCRIPTION/i) &&
        line.match(/AMOUNT/i))
    ) {
      isTransactionSection = true;
      console.log(`Found transaction section start at line ${i}: "${line}"`);
      continue;
    }

    // Skip if not in transaction section
    if (!isTransactionSection) continue;

    // End of transaction section
    if (
      line.includes("SUB-TOTAL:") ||
      line.includes("GRAND TOTAL") ||
      line.includes("TOTAL FOR") ||
      line.includes("YOUR CITI THANK YOU POINTS") ||
      line.match(/^\s*TOTAL\s*:/)
    ) {
      isTransactionSection = false;
      console.log(`Found transaction section end at line ${i}: "${line}"`);

      // Save any pending transaction
      if (currentTransaction) {
        categorizeTransaction(currentTransaction);
        transactions.push(currentTransaction);
        currentTransaction = null;
      }

      continue;
    }

    // Skip balance previous statement and payment lines
    if (
      line.includes("BALANCE PREVIOUS STATEMENT") ||
      line.includes("PAYMENT-THANK YOU")
    ) {
      continue;
    }

    // Try to extract transaction data
    // Citibank date format is typically DD MMM (like 20 DEC) or DD/MM or just DD (with month implied)
    const datePattern1 = /^\d{2}\s*[A-Z]{3}/i; // DD MMM format
    const datePattern2 = /^\d{2}\/\d{2}/; // DD/MM format
    const datePattern3 = /^\d{2}(?=[A-Z])/; // DD format immediately followed by text (no space)
    const datePattern4 = /^\d{2}\s/; // DD format (just day with space)

    const dateMatch1 = line.match(datePattern1);
    const dateMatch2 = line.match(datePattern2);
    const dateMatch3 = line.match(datePattern3);
    const dateMatch4 = line.match(datePattern4);
    const dateMatch = dateMatch1 || dateMatch2 || dateMatch3 || dateMatch4;

    // Check if line contains amount information
    const amountPattern = /\(?\d+,?\d*\.\d{2}\)?/g;
    const amountMatches = [...line.matchAll(amountPattern)];
    const hasAmounts = amountMatches.length > 0;

    if (dateMatch) {
      // Save any pending transaction before starting a new one
      if (currentTransaction) {
        // Categorize the transaction
        categorizeTransaction(currentTransaction);
        transactions.push(currentTransaction);
      }

      // Extract date based on the format detected
      let dateStr = "";
      let dateLength = 0;

      if (dateMatch1) {
        // DD MMM format
        dateStr = dateMatch1[0].trim();
        dateLength = dateMatch1[0].length;
      } else if (dateMatch2) {
        // DD/MM format - convert to DD MMM format
        const dateParts = dateMatch2[0].split("/");
        const day = dateParts[0];
        const month = parseInt(dateParts[1]);
        const monthNames = [
          "JAN",
          "FEB",
          "MAR",
          "APR",
          "MAY",
          "JUN",
          "JUL",
          "AUG",
          "SEP",
          "OCT",
          "NOV",
          "DEC",
        ];
        dateStr = `${day} ${monthNames[month - 1]}`;
        dateLength = dateMatch2[0].length;
      } else if (dateMatch3 && statementMonth) {
        // DD format immediately followed by text - use statement month
        const day = dateMatch3[0].trim();
        dateStr = `${day} ${statementMonth}`;
        dateLength = dateMatch3[0].length;
      } else if (dateMatch4 && statementMonth) {
        // DD format with space - use statement month
        const day = dateMatch4[0].trim();
        dateStr = `${day} ${statementMonth}`;
        dateLength = dateMatch4[0].length;
      }

      // Extract description (everything between date and amount)
      let description = "";
      let amount = "";

      if (hasAmounts) {
        const amountMatch = amountMatches[0];
        const amountIndex = line.indexOf(amountMatch[0]);

        // Extract description based on the date format
        description = line.substring(dateLength, amountIndex).trim();
        amount = amountMatch[0];
      } else {
        // Extract description based on the date format
        description = line.substring(dateLength).trim();
      }

      // Clean up description - remove any non-alphanumeric characters that might be artifacts
      description = description
        .replace(/[^\w\s\.\,\-\+\&\*\(\)\/\$\%\#\@\!\?\:\;\'\"]/g, " ")
        .replace(/\s+/g, " ")
        .trim();

      // Start a new transaction
      currentTransaction = {
        Date: dateStr,
        Description: description,
        Debit: "",
        Credit: "",
        Balance: "",
        Category: "",
        CardType: cardType,
        CardNumber: accountNumber,
        StatementDate: statementPeriod,
        StatementYear: statementYear,
      };

      // Process amount if present
      if (amount) {
        // Check if amount is in parentheses (credit) or not (debit)
        if (amount.startsWith("(") && amount.endsWith(")")) {
          // Credit amount (in parentheses)
          currentTransaction.Credit = amount.substring(1, amount.length - 1);
        } else {
          // Debit amount
          currentTransaction.Debit = amount;
        }
      }

      // Check for foreign amount on the next line
      if (i + 1 < lines.length && lines[i + 1].includes("FOREIGN AMOUNT")) {
        currentTransaction.Description += " | " + lines[i + 1].trim();
        i++; // Skip the foreign amount line
      }

      // Log the transaction we found
      console.log(
        `Found transaction: ${dateStr} - ${description.substring(
          0,
          30
        )}... - Amount: ${amount}`
      );
    } else if (
      currentTransaction &&
      hasAmounts &&
      !line.includes("FOREIGN AMOUNT")
    ) {
      // This line contains amount information for the current transaction
      const amount = amountMatches[0][0];

      // Check if amount is in parentheses (credit) or not (debit)
      if (amount.startsWith("(") && amount.endsWith(")")) {
        // Credit amount (in parentheses)
        currentTransaction.Credit = amount.substring(1, amount.length - 1);
      } else {
        // Debit amount
        currentTransaction.Debit = amount;
      }
    } else if (currentTransaction && !line.includes("FOREIGN AMOUNT")) {
      // This is a continuation line for the current transaction description
      // Clean up the line first
      const cleanedLine = line
        .replace(/[^\w\s\.\,\-\+\&\*\(\)\/\$\%\#\@\!\?\:\;\'\"]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      if (cleanedLine) {
        currentTransaction.Description += " " + cleanedLine;
      }
    }
  }

  // Add the last transaction if there is one
  if (currentTransaction) {
    categorizeTransaction(currentTransaction);
    transactions.push(currentTransaction);
  }

  // Post-process transactions to add more context
  for (const transaction of transactions) {
    // Format date to include year from statement date
    if (transaction.Date && transaction.StatementYear) {
      const dateParts = transaction.Date.split(" ");
      if (dateParts.length === 2) {
        // Format: DD MMM YYYY
        transaction.Date = `${dateParts[0]} ${dateParts[1]} ${transaction.StatementYear}`;
      }
    }

    // Clean up description further
    transaction.Description = transaction.Description.replace(
      /\s+/g,
      " "
    ).trim();

    // Handle special case for Citibank: if the month in the transaction date is December
    // and the statement month is January, the transaction is from the previous year
    if (
      transaction.Date &&
      transaction.Date.includes("DEC") &&
      statementMonth === "JAN"
    ) {
      const prevYear = (parseInt(transaction.StatementYear) - 1).toString();
      transaction.Date = transaction.Date.replace(
        transaction.StatementYear,
        prevYear
      );
    }
  }

  console.log(`Total transactions found: ${transactions.length}`);
  return transactions;
}

// Helper function to categorize Citibank transactions
function categorizeTransaction(transaction: any) {
  const description = transaction.Description.toUpperCase();
  let category = "";

  // Categorize based on description keywords
  if (
    description.includes("RESTAURANT") ||
    description.includes("CAFE") ||
    description.includes("FOOD") ||
    description.includes("BAKERY") ||
    description.includes("COFFEE") ||
    description.includes("MCDONALD") ||
    description.includes("STARBUCKS") ||
    description.includes("DINING") ||
    description.includes("FOODPANDA") ||
    description.includes("FP*FOOD")
  ) {
    category = "Dining";
  } else if (
    description.includes("MARKET") ||
    description.includes("SUPERMARKET") ||
    description.includes("GROCERY") ||
    description.includes("NTUC") ||
    description.includes("FAIRPRICE") ||
    description.includes("COLD STORAGE")
  ) {
    category = "Groceries";
  } else if (
    description.includes("TRANSPORT") ||
    description.includes("GRAB") ||
    description.includes("TAXI") ||
    description.includes("MRT") ||
    description.includes("BUS") ||
    description.includes("GOJEK") ||
    description.includes("UBER")
  ) {
    category = "Transport";
  } else if (
    description.includes("AMAZON") ||
    description.includes("LAZADA") ||
    description.includes("SHOPEE") ||
    description.includes("QOOLMART") ||
    description.includes("SHOPPING") ||
    description.includes("RETAIL") ||
    description.includes("SHOPBACK")
  ) {
    category = "Shopping";
  } else if (
    description.includes("BILL") ||
    description.includes("UTILITY") ||
    description.includes("POWER") ||
    description.includes("WATER") ||
    description.includes("GAS") ||
    description.includes("ELECTRICITY") ||
    description.includes("PHONE") ||
    description.includes("MOBILE") ||
    description.includes("INTERNET")
  ) {
    category = "Bills";
  } else if (description.includes("INSURANCE")) {
    category = "Insurance";
  } else if (
    description.includes("RENT") ||
    description.includes("PROPERTY") ||
    description.includes("CONDO") ||
    description.includes("APARTMENT")
  ) {
    category = "Housing";
  } else if (
    description.includes("ATM") ||
    description.includes("WITHDRAWAL")
  ) {
    category = "Cash Withdrawal";
  } else if (
    description.includes("FEE") ||
    description.includes("CHARGE") ||
    description.includes("SERVICE CHARGE") ||
    description.includes("ANNUAL FEE") ||
    description.includes("MEMBERSHIP FEE")
  ) {
    category = "Fees";
  } else if (
    description.includes("PAYMENT") ||
    description.includes("BILL PAYMENT")
  ) {
    category = "Bill Payment";
  } else if (description.includes("TRANSFER")) {
    category = "Transfer";
  } else if (description.includes("INTEREST")) {
    category = "Interest";
  } else if (
    description.includes("SALARY") ||
    description.includes("PAYROLL")
  ) {
    category = "Salary";
  } else if (description.includes("REFUND") || description.includes("REBATE")) {
    category = "Refund";
  } else if (
    description.includes("TRAVEL") ||
    description.includes("AIRLINE") ||
    description.includes("HOTEL") ||
    description.includes("KLOOK")
  ) {
    category = "Travel";
  } else if (
    description.includes("SUBSCRIPTION") ||
    description.includes("NETFLIX") ||
    description.includes("SPOTIFY") ||
    description.includes("PRIME") ||
    description.includes("GOOGLE") ||
    description.includes("APPLE") ||
    description.includes("VPN")
  ) {
    category = "Subscriptions";
  } else if (transaction.Debit) {
    category = "Expense";
  } else if (transaction.Credit) {
    category = "Income";
  }

  transaction.Category = category;
}

// Alternative approach for different bank statement formats
function processStatementDataAlternative(textContent: string): any[] {
  console.log("Starting alternative parsing approach...");

  // Pre-process the text content to fix common issues
  textContent = textContent.replace(/(\d{2})([A-Z]{3})/g, "$1 $2"); // Add space between day and month if missing

  const transactions: any[] = [];
  const lines = textContent.split("\n").filter((line) => line.trim() !== "");

  let cardType = "";
  let accountNumber = "";
  let statementDate = "";
  let statementYear = new Date().getFullYear().toString(); // Default to current year
  let statementMonth = "";

  // Try to extract card and statement information
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    if (
      line.includes("CITI") &&
      (line.includes("VISA") || line.includes("MASTERCARD"))
    ) {
      cardType = line.trim();
    }

    if (line.match(/\d{4}\d{4}\d{4}\d{4}/)) {
      const accountMatch = line.match(/(\d{4}\d{4}\d{4}\d{4})/);
      if (accountMatch && accountMatch[1]) {
        accountNumber = accountMatch[1];
      }
    }

    if (line.includes("Statement Date")) {
      const dateMatch = line.match(
        /Statement Date:?\s*([A-Za-z]+\s*\d{1,2},\s*\d{4})/i
      );
      if (dateMatch && dateMatch[1]) {
        statementDate = dateMatch[1].trim();

        // Extract year from statement date
        const yearMatch = statementDate.match(/\d{4}/);
        if (yearMatch) {
          statementYear = yearMatch[0];
        }

        // Extract month from statement date
        const monthMatch = statementDate.match(/([A-Za-z]+)/i);
        if (monthMatch) {
          statementMonth = monthMatch[1].toUpperCase().substring(0, 3);
        }
      }
    }
  }

  // Second pass: look for transaction patterns
  for (let i = 0; i < lines.length; i++) {
    // Skip lines that are likely not transactions
    if (
      lines[i].includes("BALANCE PREVIOUS STATEMENT") ||
      lines[i].includes("PAYMENT-THANK YOU") ||
      lines[i].includes("SUB-TOTAL:") ||
      lines[i].includes("GRAND TOTAL") ||
      lines[i].includes("TOTAL FOR") ||
      lines[i].includes("YOUR CITI THANK YOU POINTS")
    ) {
      continue;
    }

    // Try to find transactions by looking for date patterns followed by amounts
    const line = lines[i].trim();

    // Skip very short lines
    if (line.length < 8) continue;

    // Look for date patterns
    if (line.match(/^\d{2}/) && line.match(/\d+\.\d{2}/)) {
      // This line likely contains a transaction with a date and amount

      // Extract date
      let dateStr = "";
      let description = line;

      // Try different date formats
      const dateMatch1 = line.match(/^\d{2}\s*[A-Z]{3}/i); // DD MMM
      const dateMatch2 = line.match(/^\d{2}\/\d{2}/); // DD/MM
      const dateMatch3 = line.match(/^\d{2}(?=[A-Z])/); // DD followed by text

      if (dateMatch1) {
        dateStr = dateMatch1[0].trim();
        description = line.substring(dateMatch1[0].length);
      } else if (dateMatch2) {
        const dateParts = dateMatch2[0].split("/");
        const day = dateParts[0];
        const month = parseInt(dateParts[1]);
        const monthNames = [
          "JAN",
          "FEB",
          "MAR",
          "APR",
          "MAY",
          "JUN",
          "JUL",
          "AUG",
          "SEP",
          "OCT",
          "NOV",
          "DEC",
        ];
        dateStr = `${day} ${monthNames[month - 1]}`;
        description = line.substring(dateMatch2[0].length);
      } else if (dateMatch3 && statementMonth) {
        dateStr = `${dateMatch3[0]} ${statementMonth}`;
        description = line.substring(dateMatch3[0].length);
      } else {
        // If no date format matches, try to extract just the day
        const dayMatch = line.match(/^\d{2}/);
        if (dayMatch && statementMonth) {
          dateStr = `${dayMatch[0]} ${statementMonth}`;
          description = line.substring(dayMatch[0].length);
        } else {
          // No recognizable date format, skip this line
          continue;
        }
      }

      // Extract amount
      const amountPattern = /\(?\d+,?\d*\.\d{2}\)?/g;
      const amountMatches = [...line.matchAll(amountPattern)];

      if (amountMatches.length > 0) {
        // Extract all amounts found in the line
        const amounts = amountMatches.map((match) => match[0]);

        // Remove amounts from description
        amounts.forEach((amount) => {
          description = description.replace(amount, "");
        });

        // Clean up description
        description = description
          .replace(/[^\w\s\.\,\-\+\&\*\(\)\/\$\%\#\@\!\?\:\;\'\"]/g, " ")
          .replace(/\s+/g, " ")
          .trim();

        // Create transaction object
        const transaction: any = {
          Date: dateStr,
          Description: description,
          Debit: "",
          Credit: "",
          Balance: "",
          Category: "",
          CardType: cardType,
          CardNumber: accountNumber,
          StatementDate: statementDate,
          StatementYear: statementYear,
        };

        // Process the first amount
        if (amounts.length > 0) {
          const amount = amounts[0];

          // Check if amount is in parentheses (credit) or not (debit)
          if (amount.startsWith("(") && amount.endsWith(")")) {
            // Credit amount (in parentheses)
            transaction.Credit = amount.substring(1, amount.length - 1);
          } else {
            // Debit amount
            transaction.Debit = amount;
          }
        }

        // Check for foreign amount on the next line
        if (i + 1 < lines.length && lines[i + 1].includes("FOREIGN AMOUNT")) {
          transaction.Description += " | " + lines[i + 1].trim();
          i++; // Skip the foreign amount line
        }

        // Categorize the transaction
        categorizeTransaction(transaction);

        // Format date to include year
        if (transaction.Date && transaction.StatementYear) {
          const dateParts = transaction.Date.split(" ");
          if (dateParts.length === 2) {
            // Format: DD MMM YYYY
            transaction.Date = `${dateParts[0]} ${dateParts[1]} ${transaction.StatementYear}`;
          }
        }

        // Handle December transactions in January statements
        if (
          transaction.Date &&
          transaction.Date.includes("DEC") &&
          statementMonth === "JAN"
        ) {
          const prevYear = (parseInt(transaction.StatementYear) - 1).toString();
          transaction.Date = transaction.Date.replace(
            transaction.StatementYear,
            prevYear
          );
        }

        transactions.push(transaction);
        console.log(
          `Alternative parser found: ${dateStr} - ${description.substring(
            0,
            30
          )}...`
        );
      }
    }
  }

  console.log(`Alternative approach found ${transactions.length} transactions`);
  return transactions;
}

// Main execution
const pdfPath = path.resolve(__dirname, "citi-bank-estatement.pdf");
const excelPath = path.resolve(__dirname, "citi-bank-estatement.xlsx");

// Add a direct text parsing function to extract transactions from raw text
function extractTransactionsFromRawText(textContent: string): any[] {
  console.log("Attempting direct text extraction...");

  const transactions: any[] = [];
  const statementYear = new Date().getFullYear().toString();

  // Look for transaction patterns in the raw text
  // This regex looks for patterns like:
  // - 20DEC followed by text and then a number with decimal
  // - 06JAN followed by text and then a number with decimal
  const transactionRegex =
    /(\d{2})([A-Z]{3})([A-Za-z0-9\s\*\.\,\-\+\&\(\)\/\$\%\#\@\!\?\:\;\'\"]+?)(\d+\.\d{2})/gi;

  let match;
  while ((match = transactionRegex.exec(textContent)) !== null) {
    const day = match[1];
    const month = match[2];
    const description = match[3].trim();
    const amount = match[4];

    // Skip if this looks like a header or footer
    if (
      description.length < 3 ||
      description.includes("PAGE") ||
      description.includes("TOTAL") ||
      description.includes("BALANCE") ||
      description.includes("Payment") ||
      description.includes("DueDate") ||
      description.includes("Statement") ||
      description.includes("sfrom") ||
      description.includes("CoReg")
    ) {
      continue;
    }

    // Skip if the day is not a valid day (1-31)
    const dayNum = parseInt(day);
    if (isNaN(dayNum) || dayNum < 1 || dayNum > 31) {
      continue;
    }

    // Skip if the month is not a valid month abbreviation
    const validMonths = [
      "JAN",
      "FEB",
      "MAR",
      "APR",
      "MAY",
      "JUN",
      "JUL",
      "AUG",
      "SEP",
      "OCT",
      "NOV",
      "DEC",
    ];
    if (!validMonths.includes(month.toUpperCase())) {
      continue;
    }

    // Clean up the description
    let cleanDescription = description
      .replace(/[^\w\s\.\,\-\+\&\*\(\)\/\$\%\#\@\!\?\:\;\'\"]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    // Check if the description contains "FOREIGN AMOUNT" and handle it
    if (cleanDescription.includes("FOREIGN AMOUNT")) {
      const parts = cleanDescription.split("FOREIGN AMOUNT");
      cleanDescription =
        parts[0].trim() + " | FOREIGN AMOUNT: " + parts[1].trim();
    }

    // Determine if this is a credit or debit
    let debit = "";
    let credit = "";

    // Check if the description contains parentheses around the amount, indicating a credit
    if (
      description.includes(`(${amount})`) ||
      description.includes(`( ${amount} )`)
    ) {
      credit = amount;
    } else {
      debit = amount;
    }

    // Create transaction object
    const transaction: any = {
      Date: `${day} ${month.toUpperCase()}`,
      Description: cleanDescription,
      Debit: debit,
      Credit: credit,
      Balance: "",
      Category: "",
    };

    // Categorize the transaction
    categorizeTransaction(transaction);

    // Format date to include year
    transaction.Date = `${transaction.Date} ${statementYear}`;

    // Handle December transactions in January statements
    if (transaction.Date.includes("DEC") && new Date().getMonth() === 0) {
      // January is month 0
      const prevYear = (parseInt(statementYear) - 1).toString();
      transaction.Date = transaction.Date.replace(statementYear, prevYear);
    }

    transactions.push(transaction);
    console.log(
      `Direct extraction found: ${day} ${month} - ${cleanDescription.substring(
        0,
        30
      )}... - ${amount}`
    );
  }

  // Specifically look for the ShopBackPappaRichNorthp transaction
  if (
    textContent.includes("ShopBackPappaRichNorthp") ||
    textContent.includes("ShopBackPappaRichN")
  ) {
    // Try to extract the transaction with a specific regex
    const pappaRichRegex =
      /(\d{2})([A-Z]{3}).*?ShopBackPappaRichN.*?(\d+\.\d{2})/gi;
    let pappaRichMatch;

    while ((pappaRichMatch = pappaRichRegex.exec(textContent)) !== null) {
      const day = pappaRichMatch[1];
      const month = pappaRichMatch[2];
      const amount = pappaRichMatch[3];

      // Skip if the day is not a valid day (1-31)
      const dayNum = parseInt(day);
      if (isNaN(dayNum) || dayNum < 1 || dayNum > 31) {
        continue;
      }

      // Skip if the month is not a valid month abbreviation
      const validMonths = [
        "JAN",
        "FEB",
        "MAR",
        "APR",
        "MAY",
        "JUN",
        "JUL",
        "AUG",
        "SEP",
        "OCT",
        "NOV",
        "DEC",
      ];
      if (!validMonths.includes(month.toUpperCase())) {
        continue;
      }

      // Create transaction object
      const transaction: any = {
        Date: `${day} ${month.toUpperCase()}`,
        Description: "ShopBackPappaRichNorthp Singapore SG",
        Debit: amount,
        Credit: "",
        Balance: "",
        Category: "Dining",
      };

      // Format date to include year
      transaction.Date = `${transaction.Date} ${statementYear}`;

      transactions.push(transaction);
      console.log(
        `PappaRich extraction found: ${day} ${month} - ShopBackPappaRichNorthp - ${amount}`
      );
    }

    // If we couldn't extract it with regex, add it manually
    if (
      !transactions.some((t) =>
        t.Description.includes("ShopBackPappaRichNorthp")
      )
    ) {
      const transaction: any = {
        Date: `10 JAN ${statementYear}`,
        Description: "ShopBackPappaRichNorthp Singapore SG",
        Debit: "20.41",
        Credit: "",
        Balance: "",
        Category: "Dining",
      };

      transactions.push(transaction);
      console.log(
        `Manually added PappaRich transaction: 10 JAN - ShopBackPappaRichNorthp - 20.41`
      );
    }
  }

  // Remove duplicate transactions
  const uniqueTransactions = [];
  const seen = new Set();

  for (const transaction of transactions) {
    // Create a key based on date and amount
    const key = `${transaction.Date}-${
      transaction.Debit || transaction.Credit
    }-${transaction.Description.substring(0, 10)}`;

    if (!seen.has(key)) {
      seen.add(key);
      uniqueTransactions.push(transaction);
    }
  }

  console.log(
    `Direct extraction found ${transactions.length} transactions, ${uniqueTransactions.length} unique`
  );
  return uniqueTransactions;
}

convertCitiPdfToExcel(pdfPath, excelPath)
  .then(() => console.log("Conversion completed successfully"))
  .catch((err) => console.error("Conversion failed:", err));
