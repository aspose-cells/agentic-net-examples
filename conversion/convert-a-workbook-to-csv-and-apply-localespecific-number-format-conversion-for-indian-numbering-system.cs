// Title: Convert an Excel workbook to CSV with Indian numbering format using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to CSV with Indian number grouping | Apply Indian numbering style to numeric cells before saving as CSV using Aspose.Cells | Convert .xlsx to .csv in .NET with custom number format #,##,### | Set locale‑specific number format for all cells in Aspose.Cells and save as CSV | Use TxtSaveOptions to generate UTF‑8 CSV with Indian number format in C#
// Tags: Aspose.Cells CSV export custom number format | Indian numbering format #,##,### Aspose.Cells | apply custom style to numeric cells Aspose.Cells | TxtSaveOptions UTF-8 CSV Aspose.Cells | locale‑specific number formatting Excel to CSV .NET

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.Globalization;
using System.IO;
using System.Text;

// The example loads an XLSX workbook, iterates through every worksheet to assign the Indian numbering custom format "#,##,###" to all numeric cells, and then saves the result as a UTF‑8 encoded CSV file using Aspose.Cells TxtSaveOptions, with file existence checks and error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Indian numbering format (e.g., 1,23,456)
            const string indianNumberFormat = "#,##,###";

            // Apply the Indian number format to all numeric cells in every worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
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
                            Style style = cell.GetStyle();
                            style.Custom = indianNumberFormat;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Configure CSV save options (UTF‑8 encoding ensures proper character handling)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV)
            {
                Encoding = Encoding.UTF8
            };

            // Save the workbook as a CSV file with the applied Indian number formatting
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
