# PDF Bank Statement to Excel Converter

This script reads a PDF bank statement and exports the transaction data to an Excel file. It's specifically optimized for DBS/POSB bank statements but includes alternative parsing logic for other formats.

## Prerequisites

- Node.js (v14 or higher)
- pnpm

## Installation

1. Clone this repository or download the files
2. Install dependencies:

```bash
pnpm install
```

## Usage

1. Place your bank statement PDF file in the root directory and name it `bank-statement.pdf`
2. Run the script:

```bash
pnpm start
```

3. The script will generate an Excel file named `bank-statement.xlsx` in the same directory
4. For debugging purposes, the script also creates an `extracted-text.txt` file with the raw text content from the PDF

## Features

- Extracts transaction data including:
  - Account number
  - Transaction date
  - Detailed description with comprehensive transaction details
  - Withdrawal amount
  - Deposit amount
  - Balance
  - Transaction category
- Handles multi-page DBS/POSB bank statements
- Processes multi-line transaction entries
- Captures complete recipient information across multiple lines
- Automatically categorizes transactions based on recipient information and description
- Provides category-based summaries showing totals for each transaction type
- Preserves important transaction details like:
  - Recipient information (e.g., "TO: YOU TECHNOLOGIES GROUP (SG) PL")
  - Merchant details for debit card transactions
  - Card numbers and transaction dates
  - Reference numbers and transfer IDs
  - Company names and payment details for GIRO and Salary transactions
- Intelligently categorizes amounts as withdrawals, deposits, or balances
- Uses balance tracking to verify and correct transaction amounts
- Cleans up transaction descriptions for better readability
- Adds a summary row with total withdrawals and deposits
- Formats the Excel output with appropriate column widths
- Includes alternative parsing logic for non-DBS bank statements

## How It Works

The script uses several specialized functions to process the PDF content:

1. `processDBSStatementData`: Main parser optimized for DBS/POSB bank statements

   - Handles multi-line transaction entries
   - Extracts account numbers, dates, descriptions, and amounts
   - Collects recipient information across multiple lines until an amount appears
   - Uses context clues to categorize amounts correctly
   - Tracks balance changes to verify transaction amounts

2. `processAmounts`: Specialized function to categorize amounts

   - Uses keywords like "TO:", "FROM:", "TRANSFER", "PAYMENT" to determine transaction type
   - Handles various amount formats and positions
   - Uses balance tracking to verify withdrawal vs deposit

3. `verifyTransactionAmounts`: Ensures transaction amounts are correct

   - Uses balance changes to verify and correct withdrawal/deposit amounts
   - Resolves ambiguities in transaction categorization

4. `cleanupDescription`: Improves description readability and categorizes transactions

   - Preserves important recipient/sender information
   - Extracts merchant details for debit card transactions
   - Handles different transaction types (FAST/PAYNOW, debit card, GIRO, Salary, etc.)
   - Extracts reference numbers, card details, and transaction IDs
   - Formats descriptions for better readability
   - Assigns appropriate categories based on transaction details

5. `addSummaryRow`: Adds totals for withdrawals and deposits
   - Calculates total amounts
   - Provides category-based subtotals
   - Adds a summary row at the end of the data

## Transaction Categories

The script automatically categorizes transactions into the following types:

- **Dining**: Restaurants, cafes, food outlets, bakeries
- **Groceries**: Supermarkets, grocery stores, NTUC FairPrice, Cold Storage
- **Transport**: Grab, taxis, MRT, bus services, Gojek
- **Shopping**: Amazon, Lazada, Shopee, retail stores
- **Bills**: Utility bills, service payments
- **Telecommunications**: Phone bills, internet services
- **Utilities**: Power, water, gas bills
- **Housing**: Rent, property-related payments
- **Insurance**: Insurance premiums and payments
- **Salary**: Income from employment
- **Investment**: Securities, trading, investment platforms
- **Investment Income**: Dividends, interest from investments
- **Cash Withdrawal**: ATM withdrawals
- **Fees**: Bank fees, service charges
- **Transfer**: General fund transfers
- **Income**: Other income sources
- **Expense**: General expenses

## Transaction Type Handling

The script is optimized to handle different types of transactions:

### Debit Card Transactions

- Extracts merchant name and location
- Preserves card number information
- Captures transaction date
- Categorizes based on merchant information

### PAYNOW/FAST Transfers

- Captures recipient name (TO: field)
- Preserves transfer numbers
- Extracts reference information
- Categorizes based on recipient details

### GIRO and Salary Transactions

- Extracts company name information
- Preserves payment details and references
- Formats information in a clear, readable structure
- Categorizes based on company name and payment details

### Other Transactions

- Preserves sender information (FROM: field)
- Captures any reference numbers
- Maintains other transaction-specific details
- Assigns appropriate categories based on available information

## Customization

If you're using a different bank's statement format, you may need to customize these functions to match your specific format. The script includes an alternative parser (`processStatementDataAlternative`) that can be modified for other bank statement formats.

You can also customize the categorization logic in the `cleanupDescription` function to match your specific needs and add additional categories.

## Troubleshooting

If the script doesn't extract data correctly:

1. Check the console output for any errors
2. Examine the `extracted-text.txt` file to understand how your PDF content is structured
3. Modify the parsing functions to better match your bank statement format

## License

ISC
