// Title: C# – Set Data Label Border Thickness to 2 pt in a Stacked Bar Chart with Aspose.Cells
// Description: Learn how to create a workbook, add a stacked bar chart, enable data labels, and set the data label border weight to 2 points (with visibility) using Aspose.Cells for .NET, then save the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart data label border | set data label border thickness Aspose.Cells | stacked bar chart label styling .NET | Border.WeightPt Aspose.Cells | chart data label formatting C# | Excel chart border thickness code | Aspose.Cells chart customization
// Common Searches: Aspose.Cells set data label border weight | C# stacked bar chart data label border thickness | how to change chart label border in Aspose.Cells | Aspose.Cells chart label border visibility | set Border.WeightPt for chart data labels .NET
// Developer Intent: Set the border thickness of data labels to 2 points on a stacked bar chart using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports with stacked bar charts where data label borders highlight values. | Apply a consistent 2‑pt border to all data labels for improved readability in financial dashboards. | Programmatically enforce chart styling standards across multiple workbooks in an automated reporting pipeline.
// AI Prompts: Provide C# code that creates a stacked bar chart with Aspose.Cells and sets the data label border thickness to 2 pt. | Show how to enable and style data label borders (WeightPt and IsVisible) for a chart series in Aspose.Cells for .NET. | Explain how to apply the same 2‑point border style to every series in a stacked bar chart using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Learn how to create a workbook, add a stacked bar chart, enable data labels, and set the data label border weight to 2 points (with visibility) using Aspose.Cells for .NET, then save the file as an Excel workbook.
class SetDataLabelBorderThickness
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked bar chart
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

            // Add a stacked bar chart (use the correct ChartType enum value)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);          // values
            chart.NSeries.CategoryData = "A2:A4";      // categories

            // Access the data labels of the first series
            DataLabels labels = chart.NSeries[0].DataLabels;
            labels.ShowValue = true;                   // make data labels visible

            // Set the border thickness of the data labels to 2 points
            labels.Border.WeightPt = 2.0;
            labels.Border.IsVisible = true;            // ensure the border is shown

            // Save the workbook
            string outputPath = "StackedBarDataLabelBorder.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
