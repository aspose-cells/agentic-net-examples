using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class DeleteSecondSparkline
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two rows (A1:D2) – each row will generate a sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(3);

            sheet.Cells["A2"].PutValue(6);
            sheet.Cells["B2"].PutValue(1);
            sheet.Cells["C2"].PutValue(7);
            sheet.Cells["D2"].PutValue(4);

            // Define the location range for the sparkline group (E1:E2) – two cells for two sparklines
            CellArea location = CellArea.CreateCellArea(0, 4, 1, 4); // Rows 0‑1, Column 4 (E1:E2)

            // Add a sparkline group with the data range and location range.
            // Each row in the data range creates a sparkline in the corresponding location cell.
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Verify the number of sparklines before deletion (expected: 2)
            Console.WriteLine("Sparkline count before deletion: " + group.Sparklines.Count);

            // Delete the second sparkline (index 1) using RemoveAt
            group.Sparklines.RemoveAt(1);

            // Verify the number of sparklines after deletion (expected: 1)
            Console.WriteLine("Sparkline count after deletion: " + group.Sparklines.Count);

            // Save the workbook to a file
            string outputPath = "DeleteSecondSparkline.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}