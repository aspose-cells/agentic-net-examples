using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RemoveAllSparklines
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet (active worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Example data and sparkline creation (optional, for demonstration)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);
        CellArea sparklineLocation = new CellArea { StartColumn = 4, EndColumn = 4, StartRow = 0, EndRow = 0 };
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // ----- Core logic: remove all sparklines from the active worksheet -----
        // Iterate through each SparklineGroup in the worksheet and clear its Sparklines collection.
        foreach (SparklineGroup sg in sheet.SparklineGroups)
        {
            // Clear all sparklines within the current group.
            sg.Sparklines.Clear();
        }

        // Save the workbook to verify that sparklines have been removed.
        workbook.Save("RemovedAllSparklines.xlsx", SaveFormat.Xlsx);
    }
}