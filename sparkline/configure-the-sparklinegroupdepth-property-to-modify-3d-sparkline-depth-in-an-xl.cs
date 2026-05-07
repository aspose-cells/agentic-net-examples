using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Add data for a sparkline group (2‑D only)
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(3);
        sheet.Cells["A3"].PutValue(7);
        sheet.Cells["A4"].PutValue(2);
        sheet.Cells["A5"].PutValue(9);

        // Define where the sparkline will be placed (B1:B5)
        CellArea sparkArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 1,
            EndRow = 4,
            EndColumn = 1
        };

        // Add a line sparkline group
        int sparkGroupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, sparkArea);
        SparklineGroup sparkGroup = sheet.SparklineGroups[sparkGroupIdx];

        // -------------------------------------------------
        // Add data for a 3‑D chart (where depth can be set)
        // -------------------------------------------------
        sheet.Cells["C1"].PutValue("Category");
        sheet.Cells["C2"].PutValue("Q1");
        sheet.Cells["C3"].PutValue("Q2");
        sheet.Cells["C4"].PutValue("Q3");
        sheet.Cells["D1"].PutValue("Series1");
        sheet.Cells["D2"].PutValue(120);
        sheet.Cells["D3"].PutValue(150);
        sheet.Cells["D4"].PutValue(180);
        sheet.Cells["E1"].PutValue("Series2");
        sheet.Cells["E2"].PutValue(90);
        sheet.Cells["E3"].PutValue(110);
        sheet.Cells["E4"].PutValue(130);

        // Add a 3‑D column chart
        int chartIdx = sheet.Charts.Add(ChartType.Column3D, 6, 0, 20, 10);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("D2:E4", true);
        chart.NSeries.CategoryData = "C2:C4";

        // Configure the 3‑D depth using the DepthPercent property
        chart.DepthPercent = 150; // 150% depth

        // Save the workbook
        workbook.Save("SparklineAnd3DChartDepth.xlsx");
    }
}