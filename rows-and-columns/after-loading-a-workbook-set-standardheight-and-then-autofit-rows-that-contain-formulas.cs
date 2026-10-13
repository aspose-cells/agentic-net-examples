// Title: How to set a default row height and auto‑fit only rows containing formulas in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, assign a fixed value to Worksheet.Cells.StandardHeight, iterate through each Row to find any Cell with a non‑empty Formula, call worksheet.AutoFitRow for those rows, and save the workbook. | Write C# code that checks every row in a worksheet for formula cells, applies AutoFitRow only to rows that contain formulas after setting a standard row height, and persists the changes.
// Common Searches: Aspose.Cells set standard row height and auto fit rows with formulas in C# | C# auto fit only rows that have formulas using Aspose.Cells | detect formula cells in a worksheet and adjust row height with Aspose.Cells .NET | conditional AutoFitRow for rows containing formulas after setting default height Aspose.Cells | iterate rows and apply AutoFitRow conditionally in Aspose.Cells .NET
// Tags: set worksheet default row height Aspose.Cells | auto‑fit rows with formulas Aspose.Cells | detect formula cells in worksheet Aspose.Cells | conditional AutoFitRow usage Aspose.Cells | C# iterate rows adjust height Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, sets a fixed StandardHeight for all rows, scans each row to identify any cell containing a formula, auto‑fits only those rows using AutoFitRow, and then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine input and output file paths
            string inputPath = args.Length > 0 ? args[0] : "Input.xlsx";
            string outputPath = args.Length > 1 ? args[1] : "Output.xlsx";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook (rule)
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set the standard row height (default height for rows)
                worksheet.Cells.StandardHeight = 15; // adjust as needed

                // Auto‑fit rows that contain formulas
                foreach (Row row in worksheet.Cells.Rows)
                {
                    bool containsFormula = false;

                    foreach (Cell cell in row)
                    {
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            containsFormula = true;
                            break;
                        }
                    }

                    if (containsFormula)
                    {
                        // Auto‑fit the specific row
                        worksheet.AutoFitRow(row.Index);
                    }
                }

                // Save workbook (rule)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
