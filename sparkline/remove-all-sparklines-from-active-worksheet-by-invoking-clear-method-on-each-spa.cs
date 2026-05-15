using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklinesRemoval
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet (active worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // OPTIONAL: Add sample data and a sparkline group for demo.
            // This section can be omitted if the worksheet already contains sparklines.
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed
            CellArea sparklineLocation = new CellArea
            {
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a sparkline group (Line type) that uses the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);
            // ------------------------------------------------------------

            // Remove all sparklines from the active worksheet
            // Iterate through each SparklineGroup and clear its SparklineCollection
            foreach (SparklineGroup sg in sheet.SparklineGroups)
            {
                // Clear all sparklines within the current group
                sg.Sparklines.Clear();
            }

            // Save the workbook (lifecycle: save)
            workbook.Save("SparklinesRemoved.xlsx", SaveFormat.Xlsx);
        }
    }
}