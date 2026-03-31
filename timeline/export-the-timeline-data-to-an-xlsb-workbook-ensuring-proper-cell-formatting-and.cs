using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class ExportTimelineToXlsb
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Date and Amount columns
        cells["A1"].PutValue("Date");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue(new DateTime(2021, 1, 1));
        cells["B2"].PutValue(100);
        cells["A3"].PutValue(new DateTime(2021, 2, 1));
        cells["B3"].PutValue(200);
        cells["A4"].PutValue(new DateTime(2021, 3, 1));
        cells["B4"].PutValue(300);

        // Apply a date style to the Date column
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "m/d/yyyy";
        for (int row = 1; row <= 4; row++)
        {
            cells[row, 0].SetStyle(dateStyle);
        }

        // Add a PivotTable based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable, positioned at cell F1
        sheet.Timelines.Add(pivot, "F1", "Date");

        // Configure XLSB save options
        XlsbSaveOptions saveOptions = new XlsbSaveOptions
        {
            ExportAllColumnIndexes = true
        };

        // Save the workbook as an XLSB file
        workbook.Save("TimelineData.xlsb", saveOptions);
    }
}

class Program
{
    static void Main()
    {
        ExportTimelineToXlsb.Run();
    }
}