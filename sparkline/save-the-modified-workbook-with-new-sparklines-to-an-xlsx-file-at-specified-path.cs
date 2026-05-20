using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineSaveDemo
{
    static void Main()
    {
        // Path where the workbook with sparklines will be saved
        string outputPath = @"C:\Temp\SparklinesDemo.xlsx";

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (row 1, columns A‑D)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define the cell where the sparkline will be displayed (E1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group of type Line that uses the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,
            $"{sheet.Name}!A1:D1",
            false,
            sparklineLocation);

        // Retrieve the created group (optional: modify its properties here)
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Save the workbook to the specified path in XLSX format
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}