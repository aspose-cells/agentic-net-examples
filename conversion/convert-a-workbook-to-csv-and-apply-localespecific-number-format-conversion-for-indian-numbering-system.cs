// Title: C# – Convert Excel to CSV with Indian Number Formatting using Aspose.Cells
// Description: Loads an .xlsx workbook with the Indian culture (en‑IN), sets the region to India, defines ',' as the group separator and '.' as the decimal separator, creates a custom style "#,##,###.00" for lakh/crrore grouping, applies it to every numeric cell in the first worksheet, and saves the result as a CSV file.
// Keywords: Aspose.Cells | C# | Excel to CSV | Indian number format | en-IN culture | custom number grouping | lakh format | crrore format | SaveFormat.Csv | region India
// Common Searches: Aspose.Cells export Excel to CSV Indian format | C# convert workbook to CSV with Indian numbering | Apply '#,##,###.00' format in Aspose.Cells | Set Indian locale when saving CSV using Aspose.Cells | CSV output with lakh grouping .NET
// Developer Intent: Generate a CSV file from an Excel workbook while displaying numeric values according to the Indian numbering system.
// Use Cases: Produce CSV reports for Indian financial statements where numbers must appear in lakh/crrore format. | Create localized CSV exports for Indian users that require commas placed per Indian conventions. | Automate batch conversion of multiple workbooks to CSV while preserving Indian number formatting for downstream analytics.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel file to CSV using Indian number grouping and custom separators. | Explain how to apply the custom number format "#,##,###.00" to all numeric cells before saving as CSV with Aspose.Cells. | Suggest changes to process every worksheet in a workbook and generate separate CSV files that retain Indian formatting.

using System;
using System.Globalization;
using Aspose.Cells;

// Loads an .xlsx workbook with the Indian culture (en‑IN), sets the region to India, defines ',' as the group separator and '.' as the decimal separator, creates a custom style "#,##,###.00" for lakh/crrore grouping, applies it to every numeric cell in the first worksheet, and saves the result as a CSV file.
class ConvertWorkbookToCsvIndian
{
    static void Main()
    {
        // Paths for source workbook and destination CSV
        string sourcePath = "input.xlsx";
        string csvPath = "output.csv";

        // Load the workbook with Indian culture settings
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.CultureInfo = new CultureInfo("en-IN"); // Indian locale

        // Load the workbook using the provided LoadOptions constructor
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Apply Indian regional settings to the workbook
        workbook.Settings.Region = CountryCode.India;
        workbook.Settings.NumberGroupSeparator = ',';   // Group separator
        workbook.Settings.NumberDecimalSeparator = '.'; // Decimal separator

        // Create a style that uses Indian number grouping (e.g., 1,00,000)
        Style indianStyle = workbook.CreateStyle();
        indianStyle.Custom = "#,##,###.00";

        // Apply the Indian number format to all numeric cells in the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    cell.SetStyle(indianStyle);
                }
            }
        }

        // Save the workbook as CSV using the provided Save method and SaveFormat enum
        workbook.Save(csvPath, SaveFormat.Csv);
    }
}
