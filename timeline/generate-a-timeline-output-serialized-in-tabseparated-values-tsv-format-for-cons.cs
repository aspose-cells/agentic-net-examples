using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample date and numeric data
        sheet.Cells["A1"].PutValue("Date");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
        sheet.Cells["B4"].PutValue(200);

        // Apply date format to the date column so the pivot recognises it as a date field
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // Short date format
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        sheet.Cells.CreateRange("A2:A4").ApplyStyle(dateStyle, flag);

        // Add a PivotTable that will serve as the data source for the Timeline
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Date");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a Timeline linked to the PivotTable, positioned at cell E1
        sheet.Timelines.Add(pivot, "E1", "Date");

        // Configure TSV (tab‑separated values) save options
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Tsv)
        {
            SeparatorString = "\t"
        };

        // Save the workbook as a TSV file
        workbook.Save("TimelineOutput.tsv", saveOptions);
    }
}