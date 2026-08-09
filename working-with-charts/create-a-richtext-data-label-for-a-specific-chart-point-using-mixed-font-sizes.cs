// Title: C# – Add a Rich‑Text Data Label with Mixed Font Sizes to a Chart Point using Aspose.Cells
// Description: This example shows how to create a workbook, insert a column chart, enable a data label for the first point of the first series, and apply rich‑text formatting to the label – larger blue font for the word "High" and smaller red font for the value "10" – before saving the file as RichTextDataLabel.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# chart data label | rich text data label | mixed font size chart label | chart point custom label | DataLabels Characters method | Excel chart formatting Aspose | font color chart label | Aspose.Cells example
// Common Searches: Aspose.Cells set different font sizes in a chart data label | How to apply rich‑text to a specific chart point label in .NET | Change color of characters in Excel chart data label using Aspose | C# code for mixed‑style data label on column chart | Aspose.Cells chart point label custom formatting
// Developer Intent: Format parts of a chart point's data label with distinct font sizes and colors using Aspose.Cells for .NET.
// Use Cases: Highlight a keyword (e.g., "High") in a data label with a larger, colored font while keeping the numeric value smaller for visual emphasis. | Create multi‑style annotations on chart points to separate text and numbers within the same label. | Design threshold indicators in column charts where the label combines formatted text and values for better readability.
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and applies mixed font sizes and colors to specific characters of a data label. | Show how to use the DataLabels.Characters method to set bold, italic, and color for separate text segments of a chart point label in Aspose.Cells. | Provide an Aspose.Cells .NET snippet that formats a chart point label with "High" in 14‑pt blue and "10" in 10‑pt red.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsRichTextDataLabel
{
    // This example shows how to create a workbook, insert a column chart, enable a data label for the first point of the first series, and apply rich‑text formatting to the label – larger blue font for the word "High" and smaller red font for the value "10" – before saving the file as RichTextDataLabel.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first point of the first series
            ChartPoint point = chart.NSeries[0].Points[0];

            // Enable data label for this point and set custom text
            point.DataLabels.ShowValue = true;
            point.DataLabels.Text = "High10";

            // Apply rich‑text formatting:
            //   - Characters 0‑3 ("High") will have larger font size (14)
            //   - Characters 4‑5 ("10") will have smaller font size (10)
            point.DataLabels.Characters(0, 4).Font.Size = 14;   // "High"
            point.DataLabels.Characters(4, 2).Font.Size = 10;   // "10"

            // Optionally set different colors for demonstration
            point.DataLabels.Characters(0, 4).Font.Color = Color.Blue;
            point.DataLabels.Characters(4, 2).Font.Color = Color.Red;

            // Save the workbook
            workbook.Save("RichTextDataLabel.xlsx");
        }
    }
}
