// Title: Replace negative values with zero in the "ProfitMargins" named range using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, retrieves the named range ProfitMargins, changes any numeric cell whose value is less than 0 to 0, and saves the modified file.
// Keywords: Aspose.Cells | C# | named range | replace negative values | set negative to zero | Excel workbook | ProfitMargins | cell iteration | Workbook.Save | data sanitization
// Common Searches: Aspose.Cells replace negative numbers in named range | C# set negative values to zero in Excel | How to update cells of a named range with Aspose.Cells | Zero out negative profit margins using Aspose.Cells | Iterate through named range cells C#
// Developer Intent: Replace every numeric value lower than zero with zero inside the "ProfitMargins" named range of an Excel file.
// Use Cases: Clean profit‑margin data before financial reporting to eliminate negative entries. | Prepare spreadsheet data for charting by ensuring all margin values are non‑negative. | Sanitize imported Excel files so downstream calculations and BI tools do not encounter negative numbers.
// AI Prompts: Generate C# code with Aspose.Cells that sets any negative number to zero within a specific named range. | Explain how to handle a missing named range gracefully when updating cell values using Aspose.Cells. | Show how to log each cell that was changed from a negative value to zero in a named range. | Create a unit test that verifies negative values are replaced by zero in the ProfitMargins range.

using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceNegativeValuesInNamedRange
{
    // Loads an Excel workbook, retrieves the named range ProfitMargins, changes any numeric cell whose value is less than 0 to 0, and saves the modified file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "ProfitMargins"
                Name profitMarginsName = workbook.Worksheets.Names["ProfitMargins"];
                if (profitMarginsName == null)
                {
                    Console.WriteLine("Named range 'ProfitMargins' not found.");
                    return;
                }

                // Get the actual range object (use fully qualified name to avoid ambiguity)
                Aspose.Cells.Range profitMarginsRange = profitMarginsName.GetRange();

                // Iterate through each cell in the range
                foreach (Cell cell in profitMarginsRange)
                {
                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double currentValue = cell.DoubleValue;
                        // Replace negative numbers with zero
                        if (currentValue < 0)
                        {
                            cell.PutValue(0);
                        }
                    }
                }

                // Save the modified workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
