// Title: C# – Verify and Add a Secondary Axis in Aspose.Cells Charts (.NET)
// Description: Creates a workbook, adds a column chart with two series, checks if a secondary value axis exists using chart.HasAxis(AxisType.Value, false), adds one only when missing, then sets its title and visibility before saving.
// Keywords: Aspose.Cells | C# | .NET | secondary axis | chart.HasAxis | plot series on second axis | prevent duplicate axis | Excel chart automation | value axis | secondary value axis
// Common Searches: Aspose.Cells check if chart has secondary axis | How to add secondary axis only when not present in Aspose.Cells | chart.HasAxis secondary value axis .NET | prevent duplicate secondary axis Aspose.Cells | set secondary axis title Aspose.Cells C#
// Developer Intent: Find out how to detect an existing secondary value axis in an Aspose.Cells chart and add one only when it’s absent.
// Use Cases: Dynamic report generation where charts may already contain a secondary axis | Conditional formatting of charts based on data ranges | Ensuring clean chart layout by avoiding multiple secondary axes | Programmatically customizing secondary axis properties after confirming its presence
// AI Prompts: Write C# code using Aspose.Cells to create a chart, check chart.HasAxis(AxisType.Value, false), and if false set PlotOnSecondAxis = true and configure the secondary axis title. | Explain step‑by‑step how chart.HasAxis works for value and category axes in Aspose.Cells. | Generate a reusable method that adds a secondary axis to any Aspose.Cells chart only when it does not already exist.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart with two series, checks if a secondary value axis exists using chart.HasAxis(AxisType.Value, false), adds one only when missing, then sets its title and visibility before saving.
class CheckSecondaryAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series1");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        worksheet.Cells["C1"].PutValue("Series2");
        worksheet.Cells["C2"].PutValue(100);
        worksheet.Cells["C3"].PutValue(200);
        worksheet.Cells["C4"].PutValue(300);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add two series to the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Check whether a secondary value axis already exists
        bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);

        // If the secondary axis does not exist, enable it for the second series
        if (!hasSecondaryValueAxis)
        {
            chart.NSeries[1].PlotOnSecondAxis = true;
        }

        // Optionally customize the secondary axis after ensuring it exists
        if (chart.HasAxis(AxisType.Value, false))
        {
            chart.SecondValueAxis.Title.Text = "Secondary Axis";
            chart.SecondValueAxis.IsVisible = true;
        }

        // Save the workbook
        workbook.Save("CheckSecondaryAxis.xlsx");
    }
}
