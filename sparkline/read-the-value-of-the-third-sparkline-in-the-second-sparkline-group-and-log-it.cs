using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate data for two rows (each row will be a sparkline group)
        ws.Cells["A1"].PutValue(5);
        ws.Cells["B1"].PutValue(3);
        ws.Cells["C1"].PutValue(7);
        ws.Cells["D1"].PutValue(2);

        ws.Cells["A2"].PutValue(4);
        ws.Cells["B2"].PutValue(6);
        ws.Cells["C2"].PutValue(1);
        ws.Cells["D2"].PutValue(8);

        // Define locations where the sparklines will be placed
        CellArea locGroup1 = CellArea.CreateCellArea("E1", "E1");
        CellArea locGroup2 = CellArea.CreateCellArea("E2", "E2");

        // Add first sparkline group (row 1) and create three sparklines
        int idxGroup1 = ws.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, locGroup1);
        SparklineGroup group1 = ws.SparklineGroups[idxGroup1];
        // The group already contains one sparkline at E1; add two more at F1 and G1
        group1.Sparklines.Add(ws.Name + "!A1:D1", 0, 5); // column F (index 5)
        group1.Sparklines.Add(ws.Name + "!A1:D1", 0, 6); // column G (index 6)

        // Add second sparkline group (row 2) and create three sparklines
        int idxGroup2 = ws.SparklineGroups.Add(SparklineType.Line, "A2:D2", false, locGroup2);
        SparklineGroup group2 = ws.SparklineGroups[idxGroup2];
        // The group already contains one sparkline at E2; add two more at F2 and G2
        group2.Sparklines.Add(ws.Name + "!A2:D2", 1, 5); // column F (index 5)
        group2.Sparklines.Add(ws.Name + "!A2:D2", 1, 6); // column G (index 6)

        // Access the third sparkline (index 2) in the second sparkline group
        Sparkline thirdSparkline = group2.Sparklines[2];
        string dataRange = thirdSparkline.DataRange;

        // Log the data range
        Console.WriteLine("Third sparkline in second group DataRange: " + dataRange);

        // Save the workbook (optional)
        wb.Save("SparklinesDemo.xlsx", SaveFormat.Xlsx);
    }
}