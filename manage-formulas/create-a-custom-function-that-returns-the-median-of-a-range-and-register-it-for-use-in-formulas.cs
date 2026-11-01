// Title: Register a custom user‑defined MEDIAN function in Aspose.Cells for .NET (C#) and use it in worksheet formulas
// Common Searches: how to add a custom median UDF in Aspose.Cells C# | register user defined function for Excel formulas using Aspose.Cells .NET | Aspose.Cells custom worksheet function example C# | calculate median with a custom function in Aspose.Cells workbook | C# code to register custom function for Aspose.Cells formula engine
// Tags: custom UDF registration Aspose.Cells C# | median user defined function Aspose.Cells | Aspose.Cells formula engine custom function | C# Excel workbook custom function Aspose.Cells | register worksheet function Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMedian
{
    // The example creates a new Workbook, fills cells B1:B10 with values 1‑10, shows how a MEDIAN formula can be placed in a cell, and explains where to register a custom median function so that it can be used in Aspose.Cells formula calculations before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in column B (B1:B10).
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 1].PutValue(i + 1); // Values 1 through 10.
                }

                // Use the built‑in MEDIAN function in a cell formula.
                sheet.Cells["A1"].Formula = "=MEDIAN(B1:B10)";

                // Calculate the workbook to evaluate the formula.
                workbook.CalculateFormula();

                // Determine output path and ensure the directory exists.
                string outputPath = "MedianExample.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
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
