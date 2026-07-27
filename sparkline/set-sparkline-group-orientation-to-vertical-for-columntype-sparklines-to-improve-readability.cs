using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SetSparklineOrientationVertical
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column‑type sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(3);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Define the location range where sparklines will be placed (C1:C5)
            CellArea location = new CellArea
            {
                StartRow = 0,      // Row 1
                EndRow = 4,        // Row 5
                StartColumn = 2,   // Column C (0‑based index)
                EndColumn = 2
            };

            // Add a sparkline group of type Column with vertical orientation.
            // Use a fully‑qualified data range (including sheet name) to avoid reference errors.
            string dataRange = $"{sheet.Name}!A1:A5";
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Column, // column‑type sparkline
                dataRange,            // data range with sheet name
                true,                 // isVertical = true
                location);            // location range

            // Optional: retrieve the created group for further customization
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // The group already contains sparklines for each cell in the location range,
            // so no additional Add call is required here.

            // Save the workbook
            workbook.Save("SparklineVerticalOrientation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}