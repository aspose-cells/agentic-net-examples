using System;
using Aspose.Cells;

namespace MergedRegionsCountDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new workbook
            // If you need to load an existing file, uncomment the line below and provide the path
            // Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample merged regions for demonstration
            cells.Merge(0, 0, 2, 2); // Merge A1:B2
            cells.Merge(3, 1, 3, 3); // Merge D4:F6
            cells.Merge(7, 5, 1, 4); // Merge F8:I8

            // Get all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Calculate total number of merged regions
            int totalMergedRegions = mergedAreas.Length;

            // Output the result
            Console.WriteLine($"Total number of merged regions: {totalMergedRegions}");

            // Optionally, list each merged area
            for (int i = 0; i < mergedAreas.Length; i++)
            {
                CellArea area = mergedAreas[i];
                Console.WriteLine($"Region {i + 1}: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                                  $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");
            }

            // Save the workbook (adjust the path as needed)
            workbook.Save("MergedRegionsCountDemo.xlsx");
        }
    }
}