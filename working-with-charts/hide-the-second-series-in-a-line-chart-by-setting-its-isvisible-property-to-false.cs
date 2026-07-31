// Title: Hide a specific series in an Aspose.Cells line chart using the IsFiltered flag (C#)
// Description: Creates a workbook, adds two data series, builds a line chart, and hides the second series by setting its NSeries.IsFiltered property to true before saving the file.
// Keywords: Aspose.Cells | C# | .NET | line chart | hide chart series | IsFiltered | chart series visibility | Excel chart manipulation | filter series Aspose.Cells | NSeries.IsFiltered | hide second series
// Common Searches: Aspose.Cells hide chart series C# | How to hide a series in a line chart using Aspose.Cells | Set IsFiltered true to hide series Aspose.Cells | Remove second data series from Aspose.Cells chart | Toggle series visibility in Aspose.Cells .NET
// Developer Intent: Exclude the second data series from the rendered line chart so it does not appear in the generated Excel workbook.
// Use Cases: Display only the primary series in a multi‑series line chart while keeping other series in the data source. | Generate reports where optional series are conditionally hidden based on user settings. | Create chart templates that allow runtime toggling of series visibility via the IsFiltered flag.
// AI Prompts: Show C# code that hides a chart series in Aspose.Cells by setting NSeries.IsFiltered to true. | Explain how the IsFiltered property differs from other visibility controls in Aspose.Cells charts. | Provide an example of toggling series visibility on demand using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds two data series, builds a line chart, and hides the second series by setting its NSeries.IsFiltered property to true before saving the file.
class HideSecondSeries
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for two series
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

        // Add a line chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Add the two series to the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Hide the second series by marking it as filtered
        chart.NSeries[1].IsFiltered = true;

        // Save the workbook
        workbook.Save("HideSecondSeries.xlsx");
    }
}
