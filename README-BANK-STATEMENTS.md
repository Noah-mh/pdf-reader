# Bank Statement PDF to Excel Converter

This project provides tools to convert bank statement PDFs to Excel format, making it easier to analyze your financial data. It currently supports two banks:

1. DBS Bank statements
2. Citibank statements

## Prerequisites

- Node.js (v14 or higher)
- TypeScript
- Required npm packages:
  - pdf-parse
  - xlsx
  - fs
  - path

## Installation

1. Clone this repository
2. Install dependencies:

```bash
npm install
# or
pnpm install
```

## Usage

### Processing DBS Bank Statements

Place your DBS bank statement PDF file in the project directory with the name `bank-statement.pdf`, then run:

```bash
npx ts-node pdf-to-excel.ts
```

This will generate:

- `extracted-text.txt` - Raw text extracted from the PDF (useful for debugging)
- `bank-statement.xlsx` - Excel file with processed transactions

### Processing Citibank Statements

Place your Citibank statement PDF file in the project directory with the name `citi-bank-estatement.pdf`, then run:

```bash
pnpm start:citi
```

This will generate:

- `citi-extracted-text.txt` - Raw text extracted from the PDF (useful for debugging)
- `citi-bank-estatement.xlsx` - Excel file with processed transactions

## Features

- Extracts transaction details including:

  - Date
  - Description
  - Debit/Credit amounts
  - Balance (when available)
  - Automatically categorizes transactions based on description
  - Card/Account information (for Citibank)

- Generates summary information:
  - Total debits and credits
  - Subtotals by category

## Customization

### Adding New Categories

You can customize the transaction categorization by modifying the `categorizeTransaction` function in either script. Add new keywords to match your specific transaction descriptions.

### Supporting Other Banks

To add support for another bank:

1. Create a new TypeScript file based on one of the existing scripts
2. Modify the parsing logic to match the format of your bank's PDF statements
3. Test with sample statements to ensure accurate extraction

## Troubleshooting

If transactions are not being extracted correctly:

1. Check the extracted text file to understand how the PDF content is being parsed
2. Adjust the regular expressions in the parsing functions to better match your statement format
3. For complex statements, you may need to modify the parsing logic to handle specific formatting

## Limitations

- PDF extraction quality depends on how the PDF was generated
- Some PDFs may have security features that prevent text extraction
- Complex formatting or tables may not be parsed correctly
- Foreign currency transactions may require additional processing

## Privacy and Security

These scripts process your financial data locally on your machine. No data is sent to external servers. However, please be cautious about sharing the generated Excel files as they contain your financial information.
