// Title: Aspose.Cells C# – Disable Data Labels for the Third Series in a Column Chart
// Description: Creates a workbook, adds three data series to a column chart, enables data labels for all series, then turns off ShowValue, ShowCategoryName, and ShowPercentage for the third series only, and saves the file as an XLSX document.
// Keywords: Aspose.Cells C# example | disable data labels column chart | third series label visibility | ShowValue false Aspose.Cells | Excel chart series label control | Aspose.Cells chart customization | C# Aspose.Cells sample code | GitHub Aspose.Cells chart example | chart series hide labels | Aspose.Cells data labels property
// Common Searches: Aspose.Cells hide data labels for a specific series | C# chart series label disable Aspose.Cells | turn off third series labels column chart Aspose | ShowValue false for one series Aspose.Cells | Aspose.Cells example hide series labels
// Developer Intent: Remove data labels from only the third series of a column chart while keeping labels on the other series.
// Use Cases: Generate a sales dashboard where minor product lines are shown without label clutter. | Create a performance report that highlights primary metrics and suppresses labels for secondary data. | Export an Excel chart for presentations, omitting labels on a less important series to improve readability.
// AI Prompts: Write C# code with Aspose.Cells that disables data labels for the third series of a column chart while leaving other series labeled. | Explain how to use the DataLabels properties (ShowValue, ShowCategoryName, ShowPercentage) for individual chart series in Aspose.Cells. | Provide a conditional example that toggles label visibility for a specific series based on its index using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds three data series to a column chart, enables data labels for all series, then turns off ShowValue, ShowCategoryName, and ShowPercentage for the third series only, and saves the file as an XLSX document.
class DisableDataLabelsThirdSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: categories and three data series
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Add three series to the chart
        chart.NSeries.Add("B2:B4", true); // Series 1
        chart.NSeries.Add("C2:C4", true); // Series 2
        chart.NSeries.Add("D2:D4", true); // Series 3

        // Set category (X‑axis) data
        chart.NSeries.CategoryData = "A2:A4";

        // Optionally enable data labels for all series
        foreach (Series s in chart.NSeries)
        {
            s.DataLabels.ShowValue = true;
        }

        // Disable data labels for the third series (index 2)
        Series thirdSeries = chart.NSeries[2];
        thirdSeries.DataLabels.ShowValue = false;          // hide values
        thirdSeries.DataLabels.ShowCategoryName = false;   // hide category names
        thirdSeries.DataLabels.ShowPercentage = false;     // hide percentages

        // Save the workbook
        workbook.Save("ChartWithThirdSeriesLabelsDisabled.xlsx", SaveFormat.Xlsx);
    }
}
