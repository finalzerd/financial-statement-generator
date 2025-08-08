<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Financial Statement Generator MVP

This is a React TypeScript web application that processes Excel files containing trial balance data and generates financial statements following Thai accounting standards.

## Project Structure

- **Frontend**: React with TypeScript, Vite build tool
- **Excel Processing**: xlsx library for reading Excel files
- **File Downloads**: file-saver library for downloading generated Excel files
- **Styling**: Custom CSS with dark theme support

## Key Features

1. **File Upload**: Drag & drop Excel file upload with validation
2. **Excel Processing**: Reads trial balance data from Excel sheets
3. **Decision Tree Logic**: Implements business rules from the original VBA flowchart
4. **Financial Statement Generation**: Creates Balance Sheet, P&L, Notes, and Equity statements
5. **Download**: Exports results as Excel files

## Business Logic

The application follows the decision tree from the original Excel VBA system:
- **Sheet Validation**: Checks for 1-2 Trial Balance sheets and 1-2 Trial PL sheets
- **Processing Type**: Single-year vs Multi-year processing
- **Account Classification**: Categorizes expenses by account code ranges (5300-5311, 5312-5350, etc.)
- **Inventory Detection**: Checks for account code 1510 and purchase accounts 5010
- **Company Type**: Supports Limited Partnership (ห้างหุ้นส่วนจำกัด) and Limited Company (บริษัทจำกัด)

## Technical Implementation

- Uses TypeScript for type safety
- Modular architecture with separate services for Excel processing and statement generation
- Responsive design with CSS Grid and Flexbox
- Error handling and user feedback
- File validation and processing status indicators

## Development Notes

- The app is designed as an MVP focusing on core functionality
- Code follows the original VBA flowchart decision points
- Supports both single-year and multi-year financial statement generation
- Includes proper error handling and user experience considerations
