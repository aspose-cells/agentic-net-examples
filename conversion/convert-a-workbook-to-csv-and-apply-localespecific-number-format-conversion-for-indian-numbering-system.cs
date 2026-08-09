// Title: Aspose.Cells .NET: Convert Excel to CSV with Indian numbering format
// Description: Loads an Excel workbook, sets the region to India, defines the Indian grouping pattern "#,##,##0.00", applies the style to every numeric cell in the first worksheet, and saves the result as a CSV file.
// Keywords: Aspose.Cells | C# | .NET | Excel to CSV conversion | Indian number format | lakh crore grouping | custom number style | region India | SaveFormat.Csv | locale specific formatting
// Common Searches: Aspose.Cells export Excel as CSV with Indian grouping | C# apply '#,##,##0.00' format to all numbers before CSV export | set workbook region to India in Aspose.Cells | convert Excel to CSV using Indian number system | Aspose.Cells custom number style for India
// Developer Intent: Create a CSV file from an Excel workbook where all numeric values follow the Indian lakh‑crore grouping convention.
// Use Cases: Generate CSV reports for Indian financial statements with proper lakh/crore separators. | Prepare data extracts for Indian tax or accounting software that expects Indian number formatting. | Supply CSV feeds to dashboards that display figures in the Indian numbering style.
// AI Prompts: Show how to preserve formulas while applying the Indian number format before exporting to CSV. | Demonstrate writing the CSV to a MemoryStream instead of a file, keeping the Indian formatting intact. | Explain how to apply the Indian number style to a specific range of cells rather than the entire sheet.

using System;
using Aspose.Cells;

// Loads an Excel workbook, sets the region to India, defines the Indian grouping pattern "#,##,##0.00", applies the style to every numeric cell in the first worksheet, and saves the result as a CSV file.
class ConvertToCsvIndianNumberFormat
{
    static void Main()
    {
        // Paths for source workbook and destination CSV
        string sourcePath = "input.xlsx";
        string outputPath = "output.csv";

        // Load the workbook from the source file
        Workbook workbook = new Workbook(sourcePath);

        // Set regional settings to India to influence number formatting
        workbook.Settings.Region = CountryCode.India;
        workbook.Settings.NumberDecimalSeparator = '.';
        workbook.Settings.NumberGroupSeparator = ',';

        // Define Indian numbering format (e.g., 1,23,45,678.90)
        Style indianNumberStyle = workbook.CreateStyle();
        indianNumberStyle.Custom = "#,##,##0.00";

        // Apply the style to all numeric cells in the first worksheet
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
                    cell.SetStyle(indianNumberStyle);
                }
            }
        }

        // Save the workbook as CSV using the provided Save method
        workbook.Save(outputPath, SaveFormat.Csv);
    }
}
