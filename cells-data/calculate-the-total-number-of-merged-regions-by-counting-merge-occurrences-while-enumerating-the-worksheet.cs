using System;
using Aspose.Cells;

namespace AsposeCellsMergedRegionsCount
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample merged regions
            cells.Merge(0, 0, 2, 2); // A1:B2
            cells.Merge(3, 1, 3, 3); // B4:D6
            cells.Merge(7, 5, 1, 4); // F8:I8

            // Get all merged areas in the worksheet
            CellArea[] mergedAreas = cells.GetMergedAreas();

            // Total number of merged regions
            int totalMergedRegions = mergedAreas.Length;

            Console.WriteLine($"Total merged regions: {totalMergedRegions}");

            // Optional: display each merged area
            for (int i = 0; i < mergedAreas.Length; i++)
            {
                CellArea area = mergedAreas[i];
                Console.WriteLine($"Region {i + 1}: Rows {area.StartRow}-{area.EndRow}, Columns {area.StartColumn}-{area.EndColumn}");
            }

            // Save the workbook (optional)
            workbook.Save("MergedRegionsCount.xlsx");
        }
    }
}