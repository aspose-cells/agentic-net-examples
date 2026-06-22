using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineDynamicDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will reference
        // Data in column A (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            sheet.Cells[i, 0].PutValue(i + 1); // values 1,2,3,4,5
        }

        // Add a sparkline group of type Line
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline that uses a formula‑based data range.
        // The data range is expressed as a string; it can contain a sheet name.
        // Here we use the absolute reference "Sheet1!A1:A5".
        // The sparkline will be placed in cell B1 (row 0, column 1).
        group.Sparklines.Add(sheet.Name + "!A1:A5", 0, 1);

        // OPTIONAL: customize appearance (e.g., series color)
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = System.Drawing.Color.DarkBlue;
        group.SeriesColor = seriesColor;

        // Save the workbook with the initial sparkline
        workbook.Save("SparklineDynamicInitial.xlsx");

        // ----- Demonstrate dynamic update -----
        // Change the source data; the sparkline will reflect these changes automatically.
        sheet.Cells["A3"].PutValue(20); // modify third value
        sheet.Cells["A5"].PutValue(30); // modify fifth value

        // Recalculate formulas (not strictly required for sparklines, but ensures workbook is up‑to‑date)
        workbook.CalculateFormula();

        // Save the workbook after data change; the sparkline now shows the updated series.
        workbook.Save("SparklineDynamicUpdated.xlsx");
    }
}