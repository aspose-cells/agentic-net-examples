// Title: C# – Apply Italic Formatting to Rich‑Text Data Labels in Aspose.Cells Charts
// Description: Shows how to create a workbook, insert a column chart, enable custom data labels, set a static label text, and apply italic (optionally blue) font styling via Font.IsItalic and ApplyFont() in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart data label | italic font | rich text label | ApplyFont | column chart | custom data labels | font color | Aspose.Cells API | chart label formatting
// Common Searches: Aspose.Cells make chart data labels italic C# | set rich text formatting for Aspose.Cells chart labels | apply font style to Aspose.Cells column chart data labels | change color and italicize data labels Aspose.Cells .NET | how to use ApplyFont with chart data labels Aspose
// Developer Intent: Add italic (and optional color) styling to a chart’s rich‑text data label in an Aspose.Cells workbook using C#.
// Use Cases: Highlight key points by rendering their chart labels in italic. | Combine italic style with a distinct color for better visual emphasis in reports. | Generate presentation‑ready spreadsheets where all data labels share a uniform italic appearance.
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart and sets all data label fonts to italic and blue. | Explain how to apply different font styles (italic, bold, underline) to portions of a rich‑text data label in Aspose.Cells. | Provide a snippet that updates an existing workbook’s chart data labels to italic without altering other label settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsRichTextDataLabel
{
    // Shows how to create a workbook, insert a column chart, enable custom data labels, set a static label text, and apply italic (optionally blue) font styling via Font.IsItalic and ApplyFont() in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable data labels and customize them
            series.DataLabels.ShowValue = true;               // Show the numeric value
            series.DataLabels.IsAutoText = false;             // Allow custom text
            series.DataLabels.Text = "Custom Value";          // Set custom label text

            // Apply italic formatting to the entire data label
            series.DataLabels.Font.IsItalic = true;           // Make the label italic
            series.DataLabels.Font.Color = Color.Blue;        // Optional: change color for visibility

            // Apply the font settings to all data label instances
            series.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("RichTextDataLabelItalic.xlsx");
        }
    }
}
