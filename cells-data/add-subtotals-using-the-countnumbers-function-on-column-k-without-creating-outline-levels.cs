// Title: Insert a COUNT subtotal in column K without creating outline levels using Aspose.Cells for .NET
// Common Searches: how to insert a SUBTOTAL count formula in column K with Aspose.Cells C# without generating outline levels | Aspose.Cells C# find last non‑empty row in a specific column and add a subtotal row | add a subtotal row after data in an Excel worksheet using Aspose.Cells without enabling grouping | C# Aspose.Cells insert SUBTOTAL(2) for a range while preventing outline creation | using Aspose.Cells to calculate a count subtotal for a column and save the workbook
// Tags: insert SUBTOTAL count formula Aspose.Cells | detect last non‑empty cell column Aspose.Cells | add subtotal row without outline Aspose.Cells | Excel COUNT subtotal column K Aspose.Cells | save workbook after formula insertion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing or new workbook, determines the last populated cell in column K, inserts a new row below the data, writes a SUBTOTAL(2) formula that counts the numbers in the K column range, and saves the file—all without creating outline levels.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Load existing workbook or create a new one if the file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a workbook with one default worksheet
                }

                // Work with the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Find the last used row in column K (zero‑based column index 10).
                int lastRow = -1;
                int maxRow = sheet.Cells.MaxDataRow;
                for (int row = maxRow; row >= 0; row--)
                {
                    Cell cell = sheet.Cells[row, 10];
                    // Consider a cell non‑empty if it has a non‑null value that is not a blank string.
                    if (cell.Value != null && !(cell.Value is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        lastRow = row;
                        break;
                    }
                }

                // If data exists, insert a subtotal row below the last data row.
                if (lastRow >= 0)
                {
                    // Insert a new row after the last data row.
                    sheet.Cells.InsertRows(lastRow + 1, 1);

                    // Define the range for the SUBTOTAL function (from row 2 to the original last row).
                    int dataStartRow = 1; // zero‑based index for row 2 (assuming a header in row 1)
                    string rangeAddress = $"K{dataStartRow + 1}:K{lastRow + 1}";

                    // Place the SUBTOTAL formula (function 2 = COUNT) in the new row.
                    sheet.Cells[lastRow + 2, 10].Formula = $"SUBTOTAL(2,{rangeAddress})";
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
