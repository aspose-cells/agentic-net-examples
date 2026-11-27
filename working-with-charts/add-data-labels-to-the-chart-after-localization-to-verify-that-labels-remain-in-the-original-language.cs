// Title: Add data labels to a column chart in Aspose.Cells and confirm they retain Chinese text after switching workbook culture to French (C#)
// Common Searches: Aspose.Cells C# add data labels to column chart and keep original language after changing CultureInfo | how to preserve Chinese chart labels when setting workbook.Settings.CultureInfo to fr-FR in Aspose.Cells | verify that chart data label text remains unchanged after workbook culture change Aspose.Cells | C# example showing chart data labels before and after localization with Aspose.Cells | display and compare column chart data labels with different workbook cultures using Aspose.Cells
// Tags: Aspose.Cells add chart data labels C# | column chart data label localization Aspose.Cells | preserve chart label language after CultureInfo change | recalculate chart after workbook culture modification Aspose.Cells | save workbook with chart data labels Xlsx Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook containing Chinese category and value cells, adds a column chart with data labels that show both category names and values, prints the label texts, changes the workbook's CultureInfo to French (fr-FR), recalculates the chart, prints the label texts again to demonstrate that the labels stay in Chinese, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a workbook and fill it with sample data in a foreign language.
        // -----------------------------------------------------------------
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Header
        cells["A1"].PutValue("类别");          // Chinese for "Category"
        cells["B1"].PutValue("值");            // Chinese for "Value"

        // Data rows (Chinese fruit names)
        cells["A2"].PutValue("苹果");          // Apple
        cells["B2"].PutValue(120);
        cells["A3"].PutValue("香蕉");          // Banana
        cells["B3"].PutValue(80);
        cells["A4"].PutValue("樱桃");          // Cherry
        cells["B4"].PutValue(150);

        // -----------------------------------------------------------------
        // 2. Add a column chart and enable data labels (show category name and value).
        // -----------------------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart.
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series.
        Series series = chart.NSeries[0];
        series.DataLabels.ShowCategoryName = true;
        series.DataLabels.ShowValue = true;

        // Calculate the chart to generate the data label texts.
        chart.Calculate();

        // -----------------------------------------------------------------
        // 3. Capture and display the data label texts before any localization change.
        // -----------------------------------------------------------------
        Console.WriteLine("Data Labels BEFORE localization change:");
        PrintDataLabelTexts(series);

        // -----------------------------------------------------------------
        // 4. Change the workbook's culture to French (fr-FR) to simulate localization.
        //    This affects number/date formatting but should NOT change the language
        //    of the data labels that were taken directly from cell values.
        // -----------------------------------------------------------------
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Re‑calculate the chart after culture change.
        chart.Calculate();

        // -----------------------------------------------------------------
        // 5. Capture and display the data label texts after localization.
        // -----------------------------------------------------------------
        Console.WriteLine("\nData Labels AFTER localization change (Culture = fr-FR):");
        PrintDataLabelTexts(series);

        // -----------------------------------------------------------------
        // 6. Save the workbook for visual verification (optional).
        // -----------------------------------------------------------------
        workbook.Save("ChartDataLabelsLocalizationDemo.xlsx", SaveFormat.Xlsx);
    }

    /// <param name="series">The series whose points are inspected.</param>
    private static void PrintDataLabelTexts(Series series)
    {
        ChartPointCollection points = series.Points;
        for (int i = 0; i < points.Count; i++)
        {
            // Each point has its own DataLabels object after Calculate().
            string labelText = points[i].DataLabels?.Text ?? "(no label)";
            Console.WriteLine($"  Point {i}: {labelText}");
        }
    }
}
