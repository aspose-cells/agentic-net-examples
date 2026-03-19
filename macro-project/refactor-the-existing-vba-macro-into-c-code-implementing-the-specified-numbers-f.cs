using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

public class Program
{
    public static void Main()
    {
        NumberFormattingDemo.Run();
    }
}

public class NumberFormattingDemo
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate cells with sample data
        sheet.Cells["A1"].PutValue(1234.567);               // General number
        sheet.Cells["A2"].PutValue(0.8567);                 // Fraction
        sheet.Cells["A3"].PutValue(0.75);                   // Percentage
        sheet.Cells["A4"].PutValue(DateTime.Now);           // Date
        sheet.Cells["A5"].PutValue(TimeSpan.FromHours(1.5)); // Time

        // Copy values to column B for formatting demonstration
        sheet.Cells["B1"].PutValue(sheet.Cells["A1"].Value);
        sheet.Cells["B2"].PutValue(sheet.Cells["A2"].Value);
        sheet.Cells["B3"].PutValue(sheet.Cells["A3"].Value);
        sheet.Cells["B4"].PutValue(sheet.Cells["A4"].Value);
        sheet.Cells["B5"].PutValue(sheet.Cells["A5"].Value);

        // Apply built‑in number formats using Style.Number
        ApplyNumberFormat(sheet, "B1", 0);   // General
        ApplyNumberFormat(sheet, "B2", 2);   // Decimal with two places (0.00)
        ApplyNumberFormat(sheet, "B3", 5);   // Currency ($#,##0_);($#,##0))
        ApplyNumberFormat(sheet, "B4", 10);  // Percentage with two decimals (0.00%)
        ApplyNumberFormat(sheet, "B5", 14);  // Date (m/d/yyyy)

        // Apply a custom number format using Style.Custom and StyleFlag
        Style customStyle = workbook.CreateStyle();
        customStyle.Custom = "_-€ #,##0.00;[Red]-€ #,##0.00";
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        sheet.Cells["C1"].PutValue(9876.543);
        sheet.Cells["C1"].SetStyle(customStyle, flag);

        // ------------------------------
        // Chart with data label formatting
        // ------------------------------
        sheet.Cells["D1"].PutValue("Category");
        sheet.Cells["D2"].PutValue("A");
        sheet.Cells["D3"].PutValue("B");
        sheet.Cells["D4"].PutValue("C");
        sheet.Cells["E1"].PutValue("Value");
        sheet.Cells["E2"].PutValue(1500);
        sheet.Cells["E3"].PutValue(2500);
        sheet.Cells["E4"].PutValue(3500);

        int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("E2:E4", true);
        chart.NSeries.CategoryData = "D2:D4";

        // Enable data labels and set a custom number format
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;
        series.DataLabels.NumberFormat = "$#,##0";

        // ------------------------------
        // Pivot table with number format
        // ------------------------------
        int ptIdx = sheet.PivotTables.Add("D1:E4", "G1", "Pivot1");
        PivotTable pivot = sheet.PivotTables[ptIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        PivotField dataField = pivot.DataFields[0];
        dataField.Function = ConsolidationFunction.Sum;
        dataField.NumberFormat = "$#,##0.00";

        // Save the workbook
        workbook.Save("NumberFormattingDemo.xlsx");
    }

    // Helper method to set a built‑in number format on a cell
    private static void ApplyNumberFormat(Worksheet sheet, string cellName, int builtInNumber)
    {
        Style style = sheet.Workbook.CreateStyle();
        style.Number = builtInNumber;
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        sheet.Cells[cellName].SetStyle(style, flag);
    }
}