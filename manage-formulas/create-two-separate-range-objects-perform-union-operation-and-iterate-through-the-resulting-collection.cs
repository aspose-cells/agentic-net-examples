using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Alias to avoid conflict with System.Range (C# 8.0)
    using AsposeRange = Aspose.Cells.Range;

    public class UnionRangeDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create the first range (A1:B2) and populate it
            AsposeRange range1 = worksheet.Cells.CreateRange("A1:B2");
            range1[0, 0].PutValue("R1C1");
            range1[0, 1].PutValue("R1C2");
            range1[1, 0].PutValue("R2C1");
            range1[1, 1].PutValue("R2C2");

            // Create the second range (C1:D2) and populate it
            AsposeRange range2 = worksheet.Cells.CreateRange("C1:D2");
            range2[0, 0].PutValue("R1C3");
            range2[0, 1].PutValue("R1C4");
            range2[1, 0].PutValue("R2C3");
            range2[1, 1].PutValue("R2C4");

            // Build a UnionRange from the first range
            UnionRange unionRange = worksheet.Cells
                .CreateRange("A1:B2")
                .UnionRanges(new AsposeRange[] { range1 });

            // Add the second range to the union using the Union(string) overload
            unionRange = unionRange.Union("C1:D2");

            // Iterate through all cells in the combined UnionRange
            IEnumerator enumerator = unionRange.GetEnumerator();
            Console.WriteLine("Iterating cells in the union range:");
            while (enumerator.MoveNext())
            {
                Cell cell = enumerator.Current as Cell;
                if (cell != null)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }
            }

            // Save the workbook (optional)
            string outputPath = "UnionRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}