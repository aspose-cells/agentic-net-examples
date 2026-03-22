using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class EmbedTimelineFromMhtml
{
    static void Main()
    {
        // Create a new workbook and add sample data.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Headers
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Date");
        cells["C1"].PutValue("Value");

        // Sample data
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(new DateTime(2023, 1, 1));
        cells["C2"].PutValue(100);

        cells["A3"].PutValue("B");
        cells["B3"].PutValue(new DateTime(2023, 2, 1));
        cells["C3"].PutValue(150);

        cells["A4"].PutValue("A");
        cells["B4"].PutValue(new DateTime(2023, 3, 1));
        cells["C4"].PutValue(200);

        cells["A5"].PutValue("B");
        cells["B5"].PutValue(new DateTime(2023, 4, 1));
        cells["C5"].PutValue(250);

        // Define the data range for the PivotTable.
        string dataRange = "A1:C5";

        // Add the PivotTable starting at cell E3.
        int pivotIndex = sheet.PivotTables.Add(dataRange, "E3", "PivotTableFromMhtml");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Add fields to the PivotTable.
        pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Category
        pivot.AddFieldToArea(PivotFieldType.Column, 1);   // Date
        pivot.AddFieldToArea(PivotFieldType.Data, 2);     // Value

        // Refresh and calculate the PivotTable.
        pivot.RefreshData();
        pivot.CalculateData();

        // Save the workbook.
        string outputPath = "output_with_pivot.xlsx";
        workbook.Save(outputPath);
    }
}