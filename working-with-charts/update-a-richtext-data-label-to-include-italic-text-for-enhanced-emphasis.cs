// Title: Add Italic Formatting to Chart Data Labels with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, insert a column chart, enable data labels, set custom label text, apply italic styling via the Font.IsItalic property and ApplyFont method, and save the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart data label italic | rich text chart label | ApplyFont method | column chart example | Excel automation .NET | GitHub Aspose.Cells sample | US developers | European .NET community | India Excel charting
// Common Searches: how to make chart data labels italic using Aspose.Cells | set custom text and italic font for Excel chart labels C# | Aspose.Cells rich text formatting for chart data points | apply italic style to all data labels in a column chart | Aspose.Cells example for formatting chart labels
// Developer Intent: Apply italic styling to each chart point’s data label text.
// Use Cases: Highlight key metrics in a sales chart by displaying values in italic. | Produce financial reports that follow corporate style guidelines requiring italic data labels. | Generate automated Excel dashboards where all chart labels share a consistent italic appearance.
// AI Prompts: Generate C# code with Aspose.Cells that adds bold and italic formatting to chart data labels. | Show how to assign different font colors to individual data labels in an Aspose.Cells chart. | Explain how to combine regular and italic segments within a single chart data label using RichTextCollection.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsRichTextDataLabel
{
    // Demonstrates how to create a workbook, insert a column chart, enable data labels, set custom label text, apply italic styling via the Font.IsItalic property and ApplyFont method, and save the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["B2"].PutValue(15);
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B4"].PutValue(45);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Update each data label to include italic formatting
            foreach (ChartPoint point in series.Points)
            {
                // Set custom text for the data label
                point.DataLabels.Text = $"Value: {point.YValue}";

                // Apply italic style to the entire label
                point.DataLabels.Font.IsItalic = true;

                // Apply the font settings to the label
                point.DataLabels.ApplyFont();
            }

            // Save the workbook
            workbook.Save("RichTextDataLabelItalic.xlsx");
        }
    }
}
