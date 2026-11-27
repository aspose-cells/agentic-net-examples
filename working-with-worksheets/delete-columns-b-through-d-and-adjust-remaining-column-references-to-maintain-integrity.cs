// Title: How to delete columns B through D and keep Excel formulas intact using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete columns B to D and keep formulas updated | Aspose.Cells DeleteColumns method example with reference adjustment | Remove a range of columns in an Excel workbook and preserve formulas using .NET | How to shift column indexes after deleting multiple columns with Aspose.Cells | Update formula references automatically after column deletion in C#
// Tags: Aspose.Cells DeleteColumns update references | delete multiple worksheet columns C# | preserve formulas after column removal | adjust cell references after column deletion | worksheet column deletion .NET

using System;
using System.IO;
using Aspose.Cells;

namespace DeleteColumnsExample
{
    // The sample loads an existing workbook (or creates one with columns A‑E and a formula in E2 referencing B‑D), deletes columns B through D using worksheet.Cells.DeleteColumns(startColumnIndex, totalColumnsToDelete, true) to automatically adjust formula references, recalculates all formulas, and saves the result as output.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths
                string inputPath;
                string outputPath = "output.xlsx";

                if (args.Length >= 1 && File.Exists(args[0]))
                {
                    inputPath = args[0];
                    Console.WriteLine($"Using input file: {inputPath}");
                }
                else
                {
                    // No valid input file supplied – create a sample workbook programmatically
                    inputPath = "sample_input.xlsx";
                    Console.WriteLine("No valid input file provided. Creating a sample workbook.");

                    // Create a new workbook with sample data and a formula that references columns B‑D
                    Workbook sampleWb = new Workbook();
                    Worksheet sheet = sampleWb.Worksheets[0];

                    // Populate columns A‑E with sample values
                    sheet.Cells["A1"].PutValue("Item");
                    sheet.Cells["B1"].PutValue("Qty");
                    sheet.Cells["C1"].PutValue("Price");
                    sheet.Cells["D1"].PutValue("Discount");
                    sheet.Cells["E1"].PutValue("Total");

                    sheet.Cells["A2"].PutValue("Apple");
                    sheet.Cells["B2"].PutValue(10);
                    sheet.Cells["C2"].PutValue(2.5);
                    sheet.Cells["D2"].PutValue(0.1); // 10% discount

                    // Formula in E2: =B2*C2*(1-D2)
                    sheet.Cells["E2"].Formula = "=B2*C2*(1-D2)";

                    // Save the sample workbook to be used as input
                    sampleWb.Save(inputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Sample workbook saved as {inputPath}");
                }

                // Load the workbook (either the user‑provided one or the generated sample)
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                Console.WriteLine("Before deleting columns:");
                PrintFirstRows(worksheet, 5);

                // Delete columns B through D.
                // Column indexes are zero‑based: B = 1, D = 3.
                // Total columns to delete = 3 (B, C, D).
                int startColumnIndex = 1; // B
                int totalColumnsToDelete = 3; // B, C, D

                // The Cells.DeleteColumns method requires a third parameter to indicate whether to update references.
                worksheet.Cells.DeleteColumns(startColumnIndex, totalColumnsToDelete, true);

                // Recalculate formulas to ensure references are updated
                workbook.CalculateFormula();

                Console.WriteLine("\nAfter deleting columns B‑D:");
                PrintFirstRows(worksheet, 5);

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"\nModified workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static void PrintFirstRows(Worksheet sheet, int rowCount)
        {
            Cells cells = sheet.Cells;
            // Determine how many columns are present after deletion (up to 5 for demo)
            int maxColumns = Math.Min(5, cells.MaxColumn + 1);

            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < maxColumns; c++)
                {
                    Cell cell = cells[r, c];
                    string display = cell.IsFormula ? $"={cell.Formula}" : cell.StringValue;
                    Console.Write($"{display}\t");
                }
                Console.WriteLine();
            }
        }
    }
}
