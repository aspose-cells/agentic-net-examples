// Title: Convert an Excel workbook to CSV with Indian numbering format using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets the workbook CultureInfo to en‑IN, applies the custom pattern "#,##,##0.00" to every numeric cell, and saves the first worksheet as a CSV file with Aspose.Cells. | Demonstrate how to verify the source Excel file, create the output folder if it does not exist, iterate over each worksheet's used range to apply an Indian numeric style, and catch exceptions during the CSV export.
// Common Searches: how to export Excel to CSV with Indian number formatting using Aspose.Cells C# | apply custom '#,##,##0.00' format to all numeric cells before CSV conversion Aspose.Cells | set workbook culture to en-IN in Aspose.Cells .NET | save first worksheet as CSV while preserving locale‑specific number format | C# Aspose.Cells iterate used range to format numbers for Indian locale
// Tags: apply indian numeric format Aspose.Cells | export worksheet to csv Aspose.Cells | set workbook culture en-IN Aspose.Cells | custom number style for csv conversion .NET | iterate used range cells Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range introduced in newer C# versions
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, changes its culture to en‑IN, walks through each worksheet's used range applying the Indian numbering pattern "#,##,##0.00" to numeric cells, ensures the output directory exists, and saves the first sheet as a CSV file.
class WorkbookToCsvIndianFormat
{
    static void Main()
    {
        try
        {
            // Path to the source Excel workbook
            string sourcePath = @"C:\Input\Sample.xlsx";

            // Path for the resulting CSV file
            string csvPath = @"C:\Output\Sample_Indian.csv";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(csvPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                Console.WriteLine("Invalid output path.");
                return;
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Set the workbook culture to Indian English to affect number formatting
            workbook.Settings.CultureInfo = new CultureInfo("en-IN");

            // Define Indian number format pattern (e.g., 12,34,567.89)
            const string indianNumberPattern = "#,##,##0.00";

            // Apply the Indian number format to all numeric cells in each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null) continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Apply format only to numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            Style style = cell.GetStyle();
                            style.Custom = indianNumberPattern;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the workbook as CSV (the first worksheet will be exported)
            workbook.Save(csvPath, SaveFormat.Csv);
            Console.WriteLine($"CSV file saved to: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
