using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDataRangeAssignment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(7);
            sheet.Cells["D1"].PutValue(2);

            // Define the location where the original sparkline will be placed (cell E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with one sparkline
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, originalLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // The group already contains the first sparkline (created by Add)
            Sparkline originalSparkline = group.Sparklines[0];

            // Duplicate the sparkline:
            // Place the duplicate in cell F1 (column index 5) using the same data range as the original
            int duplicateIdx = group.Sparklines.Add(originalSparkline.DataRange, 0, 5);
            Sparkline duplicateSparkline = group.Sparklines[duplicateIdx];

            // Assign a new data range to the duplicated sparkline
            // For demonstration, use the range B1:E1
            duplicateSparkline.DataRange = "B1:E1";

            // Save the workbook to an XLSX file
            workbook.Save("DuplicatedSparklineWithNewDataRange.xlsx", SaveFormat.Xlsx);
        }
    }
}