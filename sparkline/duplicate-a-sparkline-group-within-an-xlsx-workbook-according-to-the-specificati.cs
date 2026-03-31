using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineGroupDuplicationDemo
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
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);

            // Define the location where the original sparkline will be placed (cell E1)
            CellArea originalLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add the original sparkline group
            int originalGroupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,          // type
                "A1:D1",                     // data range
                false,                       // isVertical
                originalLocation);           // location range

            SparklineGroup originalGroup = sheet.SparklineGroups[originalGroupIndex];

            // Add the sparkline to the original group (required when using Add(type) overload)
            // Here we use the overload that adds a sparkline directly with location (row, column)
            originalGroup.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ------------------------------------------------------------
            // Duplicate the sparkline group to a new location (cell F1)
            // ------------------------------------------------------------

            // Define the new location range for the duplicated sparkline (cell F1)
            CellArea duplicateLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 5, // Column F
                EndColumn = 5
            };

            // Add a new sparkline group with the same type, data range and orientation
            int duplicateGroupIndex = sheet.SparklineGroups.Add(
                originalGroup.Type,          // same sparkline type
                "A1:D1",                     // same data range
                false,                       // same orientation
                duplicateLocation);          // new location range

            SparklineGroup duplicateGroup = sheet.SparklineGroups[duplicateGroupIndex];

            // Copy each sparkline from the original group to the duplicate group
            // Adjust the column index to match the new location (offset by +1)
            foreach (Sparkline sp in originalGroup.Sparklines)
            {
                // Calculate new column based on the offset between original and duplicate locations
                int columnOffset = duplicateLocation.StartColumn - originalLocation.StartColumn;
                int newColumn = sp.Column + columnOffset;

                // Add the sparkline to the duplicate group with the same data range
                duplicateGroup.Sparklines.Add(sp.DataRange, sp.Row, newColumn);
            }

            // Save the workbook (output file)
            workbook.Save("DuplicatedSparklineGroup.xlsx");
        }
    }
}