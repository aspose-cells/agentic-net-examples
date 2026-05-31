using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class NonContiguousNamedRangeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (default name is "Sheet1")
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in non‑contiguous cells
                cells["A1"].PutValue(10);
                cells["C3"].PutValue(20);
                cells["E5"].PutValue(30);

                // Add a named range that refers to the non‑contiguous cells
                int nameIndex = workbook.Worksheets.Names.Add("MyNonContig");
                Name namedRange = workbook.Worksheets.Names[nameIndex];

                // Assign a custom formula to the name (sum of the three cells)
                namedRange.RefersTo = "=SUM(Sheet1!$A$1,$C$3,$E$5)";

                // Use the named range in a worksheet formula
                cells["G1"].Formula = "=MyNonContig";

                // Calculate all formulas so that G1 shows the result
                workbook.CalculateFormula();

                // Save the workbook
                string outputPath = "NonContiguousNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            NonContiguousNamedRangeDemo.Run();
        }
    }
}