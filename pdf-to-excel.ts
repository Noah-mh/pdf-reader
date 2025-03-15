import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import pdfParse from 'pdf-parse';

async function convertPdfToExcel(pdfPath: string, excelPath: string): Promise<void> {
  try {
    console.log(`Reading PDF file: ${pdfPath}`);
    
    // Read the PDF file
    const pdfBuffer = fs.readFileSync(pdfPath);
    const pdfData = await pdfParse(pdfBuffer);
    
    // Extract text content
    const textContent = pdfData.text;
    console.log('PDF content extracted successfully');
    
    // Save the extracted text to a file for debugging
    fs.writeFileSync('extracted-text.txt', textContent);
    console.log('Extracted text saved to extracted-text.txt for debugging');
    
    // Process the text content to extract structured data
    const data = processDBSStatementData(textContent);
    
    console.log(`Found ${data.length} transactions`);
    
    // If no transactions were found, try a more generic approach
    if (data.length === 0) {
      console.log('No transactions found with standard parsing. Trying alternative approach...');
      const alternativeData = processStatementDataAlternative(textContent);
      console.log(`Alternative approach found ${alternativeData.length} transactions`);
      
      if (alternativeData.length > 0) {
        console.log('Using data from alternative parsing approach');
        data.push(...alternativeData);
      }
    }
    
    // If still no data, add a dummy row to show the file isn't completely empty
    if (data.length === 0) {
      data.push({
        Note: 'No transactions could be automatically extracted from the PDF. Please check extracted-text.txt and adjust the parsing logic.'
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
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Bank Statement');
    
    // Write to Excel file
    XLSX.writeFile(workbook, excelPath);
    
    console.log(`Excel file created successfully: ${excelPath}`);
  } catch (error) {
    console.error('Error converting PDF to Excel:', error);
    throw error;
  }
}

// Helper function to add a summary row with totals
function addSummaryRow(data: any[]) {
  // Calculate totals
  let totalWithdrawal = 0;
  let totalDeposit = 0;
  
  for (const transaction of data) {
    // Convert withdrawal and deposit to numbers
    if (transaction.Withdrawal) {
      const withdrawal = parseFloat(transaction.Withdrawal.replace(/,/g, ''));
      if (!isNaN(withdrawal)) {
        totalWithdrawal += withdrawal;
      }
    }
    
    if (transaction.Deposit) {
      const deposit = parseFloat(transaction.Deposit.replace(/,/g, ''));
      if (!isNaN(deposit)) {
        totalDeposit += deposit;
      }
    }
  }
  
  // Add summary row
  data.push({
    Date: '',
    Description: 'TOTAL',
    Withdrawal: totalWithdrawal.toFixed(2),
    Deposit: totalDeposit.toFixed(2),
    Balance: '',
    Account: ''
  });
}

// Helper function to format the worksheet
function formatWorksheet(worksheet: XLSX.WorkSheet, data: any[]) {
  // Set column widths
  const columnWidths = [
    { wch: 12 },  // Date
    { wch: 50 },  // Description
    { wch: 15 },  // Withdrawal
    { wch: 15 },  // Deposit
    { wch: 15 },  // Balance
    { wch: 15 }   // Account
  ];
  
  worksheet['!cols'] = columnWidths;
}

// Specific parser for DBS bank statements
function processDBSStatementData(textContent: string): any[] {
  // Split the content by lines
  const lines = textContent.split('\n').filter(line => line.trim() !== '');
  
  // Array to store the extracted transactions
  const transactions: any[] = [];
  
  let isTransactionSection = false;
  let currentAccount = '';
  let currentTransaction: any = null;
  let multiLineDescription = '';
  let lastBalance = 0; // Track the last known balance
  let recipientLines = []; // Track recipient information lines
  let isCollectingRecipientInfo = false; // Flag to indicate we're collecting recipient info
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Detect account number
    if (line.includes('Account No.')) {
      const accountMatch = line.match(/Account No\.\s+(\d+-\d+-\d+)/);
      if (accountMatch && accountMatch[1]) {
        currentAccount = accountMatch[1];
        console.log(`Found account number: ${currentAccount}`);
      }
    }
    
    // Detect the start of transaction section
    if (line.includes('Transaction Details') || 
        (line.includes('Date') && line.includes('Description') && line.includes('Withdrawal') && line.includes('Deposit') && line.includes('Balance'))) {
      isTransactionSection = true;
      console.log(`Found transaction section start at line ${i}: "${line}"`);
      continue;
    }
    
    // Skip if not in transaction section
    if (!isTransactionSection) continue;
    
    // Check for balance brought forward to initialize lastBalance
    if (line.includes('Balance Brought Forward')) {
      const balanceMatch = line.match(/SGD\s+([\d,]+\.\d{2})/);
      if (balanceMatch && balanceMatch[1]) {
        lastBalance = parseFloat(balanceMatch[1].replace(/,/g, ''));
        console.log(`Initial balance: ${lastBalance}`);
      }
      console.log(`Skipping header line: "${line}"`);
      continue;
    }
    
    // Skip headers or summary lines
    if (line.includes('CURRENCY:') || line.includes('Account Summary')) {
      console.log(`Skipping header line: "${line}"`);
      continue;
    }
    
    // End of transaction section
    if (line.includes('Total') || line.includes('End of Statement')) {
      isTransactionSection = false;
      console.log(`Found transaction section end at line ${i}: "${line}"`);
      
      // Save any pending transaction
      if (currentTransaction) {
        // Add recipient information if available
        if (recipientLines.length > 0) {
          currentTransaction.RecipientInfo = recipientLines.join(' ');
        }
        
        // Verify the transaction amounts using balance
        verifyTransactionAmounts(currentTransaction, lastBalance);
        
        // Update lastBalance if this transaction has a balance
        if (currentTransaction.Balance) {
          lastBalance = parseFloat(currentTransaction.Balance.replace(/,/g, ''));
        }
        
        transactions.push(currentTransaction);
        currentTransaction = null;
        multiLineDescription = '';
        recipientLines = [];
        isCollectingRecipientInfo = false;
      }
      
      continue;
    }
    
    // Try to extract transaction data
    // DBS date format is DD/MM/YYYY
    const datePattern = /\d{2}\/\d{2}\/\d{4}/;
    const dateMatch = line.match(datePattern);
    
    // Check if line contains amount information
    const amountPattern = /\d+,?\d*\.\d{2}/g;
    const amountMatches = [...line.matchAll(amountPattern)];
    const hasAmounts = amountMatches.length > 0;
    
    if (dateMatch) {
      // Save any pending transaction before starting a new one
      if (currentTransaction) {
        // Add recipient information if available
        if (recipientLines.length > 0) {
          currentTransaction.RecipientInfo = recipientLines.join(' ');
        }
        
        // Verify the transaction amounts using balance
        verifyTransactionAmounts(currentTransaction, lastBalance);
        
        // Update lastBalance if this transaction has a balance
        if (currentTransaction.Balance) {
          lastBalance = parseFloat(currentTransaction.Balance.replace(/,/g, ''));
        }
        
        transactions.push(currentTransaction);
      }
      
      const date = dateMatch[0];
      
      // Extract the rest of the line after the date
      let remainingText = line.substring(line.indexOf(date) + date.length).trim();
      
      // Start a new transaction
      currentTransaction = {
        Account: currentAccount,
        Date: date,
        Description: remainingText,
        Withdrawal: '',
        Deposit: '',
        Balance: '',
        RecipientInfo: ''
      };
      
      multiLineDescription = remainingText;
      recipientLines = []; // Reset recipient lines for new transaction
      
      // Start collecting recipient info if this is a transaction with recipient details
      isCollectingRecipientInfo = remainingText.includes('Advice') || 
                                 remainingText.includes('FAST') || 
                                 remainingText.includes('TRANSFER') || 
                                 remainingText.includes('Debit Card') ||
                                 remainingText.includes('GIRO') ||
                                 remainingText.includes('Salary');
      
      // Add the first line to recipient info if it contains relevant information
      if (isCollectingRecipientInfo) {
        recipientLines.push(remainingText);
      }
      
      // Process amounts if present
      if (hasAmounts) {
        processAmounts(currentTransaction, line, amountMatches, lastBalance);
        isCollectingRecipientInfo = false; // Stop collecting recipient info if we found amounts
      }
    } else if (currentTransaction) {
      // This is a continuation line for the current transaction
      
      if (hasAmounts) {
        // This line contains amount information - process it and stop collecting recipient info
        processAmounts(currentTransaction, line, amountMatches, lastBalance);
        isCollectingRecipientInfo = false;
      } else {
        // This is part of the description or recipient info
        multiLineDescription += ' ' + line;
        currentTransaction.Description = multiLineDescription.replace(/\s+/g, ' ').trim();
        
        // If we're collecting recipient info and there are no amounts, add this line
        if (isCollectingRecipientInfo) {
          recipientLines.push(line);
        } else if (!line.match(/^\s*$/)) {
          // Even if we're not explicitly collecting recipient info, 
          // still capture non-empty continuation lines as they might contain important details
          recipientLines.push(line);
        }
      }
    }
  }
  
  // Add the last transaction if there is one
  if (currentTransaction) {
    // Add recipient information if available
    if (recipientLines.length > 0) {
      currentTransaction.RecipientInfo = recipientLines.join(' ');
    }
    
    // Verify the transaction amounts using balance
    verifyTransactionAmounts(currentTransaction, lastBalance);
    transactions.push(currentTransaction);
  }
  
  // Clean up descriptions
  for (const transaction of transactions) {
    cleanupDescription(transaction);
  }
  
  return transactions;
}

// Helper function to process amounts in a line
function processAmounts(transaction: any, line: string, amountMatches: RegExpMatchArray[], lastBalance: number = 0) {
  // Extract all amounts
  const amounts = amountMatches.map(match => match[0]);
  
  // For DBS statements, we need to determine if this is a withdrawal or deposit
  // based on context clues in the description and surrounding text
  
  // Check for withdrawal indicators
  let isWithdrawal = 
    line.includes('TO:') || 
    (line.includes('TRANSFER') && !line.includes('FROM:') && !line.includes('INCOMING')) || 
    line.includes('PAYMENT') && !line.includes('INCOMING') || 
    line.includes('Debit Card') ||
    line.includes('PURCHASE') ||
    line.includes('ATM') ||
    line.includes('WITHDRAWAL');
  
  // Check for deposit indicators
  let isDeposit = 
    line.includes('FROM:') || 
    line.includes('INCOMING') || 
    line.includes('RECEIPT') || 
    line.includes('SALARY') || 
    line.includes('INTEREST') ||
    line.includes('CREDIT') ||
    line.includes('REFUND');
  
  // If transaction description contains both withdrawal and deposit indicators,
  // we need to check the full context of the transaction
  if (isWithdrawal && isDeposit) {
    // In case of conflict, check the full transaction description
    const fullDesc = transaction.Description.toUpperCase();
    
    // If the description has more withdrawal indicators than deposit indicators,
    // treat it as a withdrawal
    const withdrawalScore = 
      (fullDesc.includes('TO:') ? 1 : 0) + 
      (fullDesc.includes('TRANSFER') && !fullDesc.includes('FROM:') ? 1 : 0) +
      (fullDesc.includes('PAYMENT') && !fullDesc.includes('INCOMING') ? 1 : 0) +
      (fullDesc.includes('DEBIT') ? 1 : 0);
    
    const depositScore = 
      (fullDesc.includes('FROM:') ? 1 : 0) + 
      (fullDesc.includes('INCOMING') ? 1 : 0) +
      (fullDesc.includes('RECEIPT') ? 1 : 0) +
      (fullDesc.includes('SALARY') ? 1 : 0) +
      (fullDesc.includes('CREDIT') ? 1 : 0);
    
    if (withdrawalScore > depositScore) {
      // This is likely a withdrawal
      isWithdrawal = true;
      isDeposit = false;
    } else if (depositScore > withdrawalScore) {
      // This is likely a deposit
      isWithdrawal = false;
      isDeposit = true;
    }
    // If scores are equal, we'll rely on the amount pattern
  }
  
  // For DBS statements, the pattern is typically:
  // For withdrawals: [amount] [balance]
  // For deposits: [amount] [balance]
  // We need to determine which is which based on context
  
  if (amounts.length === 1) {
    // Only one amount found - could be withdrawal, deposit, or balance
    // For DBS, if there's only one amount on a line with a date, it's usually the balance
    // But if it's a continuation line, it could be the transaction amount
    if (line.match(/\d{2}\/\d{2}\/\d{4}/)) {
      // Line contains a date - the amount is likely the balance
      transaction.Balance = amounts[0];
    } else if (isWithdrawal) {
      transaction.Withdrawal = amounts[0];
    } else if (isDeposit) {
      transaction.Deposit = amounts[0];
    } else {
      // If we can't determine, assume it's the balance
      transaction.Balance = amounts[0];
    }
  } else if (amounts.length === 2) {
    // Two amounts - typically transaction amount and balance
    // For DBS, the first amount is usually the transaction amount, second is balance
    if (isWithdrawal) {
      transaction.Withdrawal = amounts[0];
      transaction.Balance = amounts[1];
    } else if (isDeposit) {
      transaction.Deposit = amounts[0];
      transaction.Balance = amounts[1];
    } else {
      // If we can't determine from context, check if balance is decreasing
      const currentAmount = parseFloat(amounts[0].replace(/,/g, ''));
      const balanceAmount = parseFloat(amounts[1].replace(/,/g, ''));
      
      if (lastBalance > 0 && balanceAmount < lastBalance) {
        // Balance decreased - this is a withdrawal
        transaction.Withdrawal = amounts[0];
        transaction.Balance = amounts[1];
      } else if (lastBalance > 0 && balanceAmount > lastBalance) {
        // Balance increased - this is a deposit
        transaction.Deposit = amounts[0];
        transaction.Balance = amounts[1];
      } else {
        // Can't determine from balance change, use position
        // In DBS format, first amount is transaction, second is balance
        transaction.Balance = amounts[1];
        
        // Use context clues from description to determine if withdrawal or deposit
        if (transaction.Description.includes('TO:') || 
            transaction.Description.includes('PAYMENT') || 
            transaction.Description.includes('Debit Card')) {
          transaction.Withdrawal = amounts[0];
        } else {
          transaction.Deposit = amounts[0];
        }
      }
    }
  } else if (amounts.length >= 3) {
    // Three or more amounts - could be multiple transactions or complex format
    // For DBS, if there are 3 amounts on a line with transaction details,
    // they are often: withdrawal, deposit, balance
    
    // Check if any amount is already set
    if (!transaction.Withdrawal && !transaction.Deposit) {
      // No amounts set yet - assign based on context
      if (isWithdrawal) {
        transaction.Withdrawal = amounts[0];
        transaction.Balance = amounts[amounts.length - 1];
      } else if (isDeposit) {
        transaction.Deposit = amounts[0];
        transaction.Balance = amounts[amounts.length - 1];
      } else {
        // If we can't determine, use position
        // In DBS format, withdrawal is usually first, deposit second, balance last
        if (parseFloat(amounts[0].replace(/,/g, '')) > 0) {
          if (line.includes('(-)')) {
            transaction.Withdrawal = amounts[0];
          } else {
            transaction.Deposit = amounts[0];
          }
        }
        transaction.Balance = amounts[amounts.length - 1];
      }
    } else {
      // Some amounts already set - just update the balance
      transaction.Balance = amounts[amounts.length - 1];
    }
  }
  
  // Final check - if we have both withdrawal and deposit set, keep only the one with context support
  if (transaction.Withdrawal && transaction.Deposit) {
    if (isWithdrawal && !isDeposit) {
      transaction.Deposit = '';
    } else if (isDeposit && !isWithdrawal) {
      transaction.Withdrawal = '';
    }
    // If both have context support, keep both (split transaction)
  }
}

// Helper function to verify transaction amounts using balance
function verifyTransactionAmounts(transaction: any, lastBalance: number) {
  // If we have a balance, we can verify the withdrawal/deposit
  if (transaction.Balance && lastBalance > 0) {
    const currentBalance = parseFloat(transaction.Balance.replace(/,/g, ''));
    const balanceDifference = currentBalance - lastBalance;
    
    // If balance decreased, this should be a withdrawal
    if (balanceDifference < 0) {
      const expectedWithdrawal = Math.abs(balanceDifference).toFixed(2);
      
      // If withdrawal is not set or doesn't match expected, update it
      if (!transaction.Withdrawal || 
          parseFloat(transaction.Withdrawal.replace(/,/g, '')) !== parseFloat(expectedWithdrawal)) {
        
        // Only update if the difference is significant (more than 0.01)
        if (Math.abs(parseFloat(transaction.Withdrawal.replace(/,/g, '')) - parseFloat(expectedWithdrawal)) > 0.01) {
          // Clear any incorrect deposit
          transaction.Deposit = '';
          transaction.Withdrawal = expectedWithdrawal;
        }
      }
    } 
    // If balance increased, this should be a deposit
    else if (balanceDifference > 0) {
      const expectedDeposit = balanceDifference.toFixed(2);
      
      // If deposit is not set or doesn't match expected, update it
      if (!transaction.Deposit || 
          parseFloat(transaction.Deposit.replace(/,/g, '')) !== parseFloat(expectedDeposit)) {
        
        // Only update if the difference is significant (more than 0.01)
        if (Math.abs(parseFloat(transaction.Deposit.replace(/,/g, '')) - parseFloat(expectedDeposit)) > 0.01) {
          // Clear any incorrect withdrawal
          transaction.Withdrawal = '';
          transaction.Deposit = expectedDeposit;
        }
      }
    }
  }
}

// Helper function to clean up transaction descriptions
function cleanupDescription(transaction: any) {
  let description = transaction.Description;
  
  // Remove any amount patterns from the description
  const amountPattern = /\d+,?\d*\.\d{2}/g;
  description = description.replace(amountPattern, '');
  
  // Remove any withdrawal/deposit indicators
  description = description.replace(/\(\-\)|\(\+\)/g, '');
  
  // If the description contains "Advice", try to extract a cleaner description
  if (description.includes('Advice')) {
    // For DBS statements, the actual transaction description often follows "Advice"
    const parts = description.split(/\s+/);
    if (parts.length > 2) {
      // Skip "Advice" and the next word (like "FAST")
      description = parts.slice(2).join(' ');
    }
  }
  
  // Clean up the description further
  description = description.replace(/\s+/g, ' ').trim();
  
  // Use the collected recipient information if available
  if (transaction.RecipientInfo) {
    // Extract the most important parts from recipient info
    let recipientInfo = transaction.RecipientInfo;
    
    // For Debit Card transactions, extract merchant and location
    if (recipientInfo.includes('Debit Card Transaction')) {
      // Extract merchant name and location
      // The pattern is typically: "Debit Card Transaction MERCHANT_NAME LOCATION DATE"
      const merchantMatch = recipientInfo.match(/Debit Card Transaction\s+(.+?)(?:\s+\d{2}[A-Z]{3}|\s+\d{4}-\d{4}|$)/);
      if (merchantMatch && merchantMatch[1]) {
        const merchantInfo = merchantMatch[1].trim();
        
        // Try to separate merchant name and location if possible
        const merchantParts = merchantInfo.split(/\s+(?=[A-Z]{3}$)/);
        if (merchantParts.length > 1) {
          description = `${description} | MERCHANT: ${merchantParts[0].trim()} | LOCATION: ${merchantParts[1].trim()}`;
        } else {
          description = `${description} | MERCHANT: ${merchantInfo}`;
        }
      }
      
      // Extract card number if present
      const cardMatch = recipientInfo.match(/(\d{4}-\d{4}-\d{4}-\d{4})/);
      if (cardMatch && cardMatch[1]) {
        description = `${description} | CARD: ${cardMatch[1]}`;
      }
      
      // Extract transaction date if present (typically in format "DDM" like "14JAN")
      const dateMatch = recipientInfo.match(/\s+(\d{2}[A-Z]{3})\s*$/);
      if (dateMatch && dateMatch[1]) {
        description = `${description} | TRANSACTION DATE: ${dateMatch[1]}`;
      }
    } else {
      // For PAYNOW/FAST transactions
      if (recipientInfo.includes('PAYNOW') || recipientInfo.includes('FAST')) {
        // Extract transfer number
        const transferMatch = recipientInfo.match(/TRANSFER\s+(\d+)/);
        let transferInfo = '';
        if (transferMatch && transferMatch[1]) {
          transferInfo = `TRANSFER: ${transferMatch[1].trim()}`;
        }
        
        // Extract TO: or FROM: information
        let recipientName = '';
        const toMatch = recipientInfo.match(/TO:\s+([^\n]+)/);
        const fromMatch = recipientInfo.match(/FROM:\s+([^\n]+)/);
        
        if (toMatch && toMatch[1]) {
          recipientName = `TO: ${toMatch[1].trim()}`;
        } else if (fromMatch && fromMatch[1]) {
          recipientName = `FROM: ${fromMatch[1].trim()}`;
        }
        
        // Extract reference numbers
        let refInfo = '';
        const refMatch = recipientInfo.match(/REF\s+([^\s]+)/);
        if (refMatch && refMatch[1]) {
          refInfo = `REF: ${refMatch[1].trim()}`;
        }
        
        // Combine the information
        if (recipientName) {
          description = `${description} | ${recipientName}`;
        }
        
        if (transferInfo) {
          description = `${description} | ${transferInfo}`;
        }
        
        if (refInfo) {
          description = `${description} | ${refInfo}`;
        }
      } else {
        // For other types of transactions
        // Extract TO: or FROM: information
        let toFromInfo = '';
        const toMatch = recipientInfo.match(/TO:\s+([^\n]+)/);
        const fromMatch = recipientInfo.match(/FROM:\s+([^\n]+)/);
        
        if (toMatch && toMatch[1]) {
          toFromInfo = `TO: ${toMatch[1].trim()}`;
        } else if (fromMatch && fromMatch[1]) {
          toFromInfo = `FROM: ${fromMatch[1].trim()}`;
        }
        
        // Handle GIRO and Salary transactions
        if (recipientInfo.includes('GIRO') || recipientInfo.includes('Salary')) {
          // For GIRO/Salary, the company name is often on the second line
          const lines: string[] = recipientInfo.split('\n');
          const filteredLines: string[] = [];
          
          for (const line of lines) {
            const trimmedLine = line.trim();
            if (trimmedLine) {
              filteredLines.push(trimmedLine);
            }
          }
          
          // For GIRO Salary, the format is typically:
          // Line 1: "GIRO Salary"
          // Line 2: "COMPANY NAME"
          // Line 3: "PAYMENT DETAILS"
          
          // Extract the main description (first line)
          const mainDesc = filteredLines[0].trim();
          
          // Reset the description to just the transaction type
          if (mainDesc.includes('GIRO') || mainDesc.includes('Salary')) {
            const transactionType = mainDesc.includes('GIRO') ? 'GIRO' : 'Salary';
            description = transactionType;
          }
          
          if (filteredLines.length >= 2) {
            // The second line typically contains the company name
            const companyName = filteredLines[1].trim();
            if (companyName) {
              description = `${description} | FROM: ${companyName}`;
            }
            
            // If there's a third line, it often contains payment details
            if (filteredLines.length >= 3) {
              const paymentDetails = filteredLines[2].trim();
              if (paymentDetails) {
                description = `${description} | DETAILS: ${paymentDetails}`;
              }
            }
          }
        }
        
        // Extract reference numbers
        let refInfo = '';
        const refMatch = recipientInfo.match(/REF\s+([^\s]+)/);
        if (refMatch && refMatch[1]) {
          refInfo = `REF: ${refMatch[1].trim()}`;
        }
        
        // Extract transaction numbers
        const transactionMatch = recipientInfo.match(/TRANSFER\s+(\d+)/);
        if (transactionMatch && transactionMatch[1]) {
          if (refInfo) {
            refInfo += ` | TRANSFER: ${transactionMatch[1].trim()}`;
          } else {
            refInfo = `TRANSFER: ${transactionMatch[1].trim()}`;
          }
        }
        
        // Combine the information
        if (toFromInfo) {
          description = `${description} | ${toFromInfo}`;
        }
        
        if (refInfo) {
          description = `${description} | ${refInfo}`;
        }
      }
    }
  }
  
  // Update the transaction description
  transaction.Description = description;
}

// Alternative approach for different bank statement formats
function processStatementDataAlternative(textContent: string): any[] {
  const transactions: any[] = [];
  const lines = textContent.split('\n').filter(line => line.trim() !== '');
  
  // Look for patterns that might indicate a transaction
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Skip very short lines
    if (line.length < 10) continue;
    
    // Look for lines with dates in DD/MM/YYYY format
    const datePattern = /\d{2}\/\d{2}\/\d{4}/;
    const dateMatch = line.match(datePattern);
    
    if (dateMatch) {
      const date = dateMatch[0];
      
      // Look for amounts in the format of numbers with decimal points
      const amountPattern = /\d+,?\d*\.\d{2}/g;
      const amountMatches = [...line.matchAll(amountPattern)];
      
      if (amountMatches.length > 0) {
        // Extract all amounts found in the line
        const amounts = amountMatches.map(match => match[0]);
        
        // Get description by removing date and amounts
        let description = line.replace(datePattern, '').trim();
        amounts.forEach(amount => {
          description = description.replace(amount, '');
        });
        
        // Clean up the description
        description = description.replace(/\s+/g, ' ').trim();
        
        // Create transaction object with different fields depending on number of amounts
        const transaction: any = {
          Date: date,
          Description: description
        };
        
        if (amounts.length >= 3) {
          transaction.Withdrawal = amounts[0];
          transaction.Deposit = amounts[1];
          transaction.Balance = amounts[2];
        } else if (amounts.length === 2) {
          // Assume first is transaction amount, second is balance
          transaction.Amount = amounts[0];
          transaction.Balance = amounts[1];
        } else if (amounts.length === 1) {
          transaction.Amount = amounts[0];
        }
        
        transactions.push(transaction);
        
        console.log(`Alternative parser found: Date=${date}, Description=${description.substring(0, 30)}...`);
      }
    }
  }
  
  return transactions;
}

// Main execution
const pdfPath = path.resolve(__dirname, 'bank-statement.pdf');
const excelPath = path.resolve(__dirname, 'bank-statement.xlsx');

convertPdfToExcel(pdfPath, excelPath)
  .then(() => console.log('Conversion completed successfully'))
  .catch(err => console.error('Conversion failed:', err)); 