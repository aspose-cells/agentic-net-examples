// Title: C# – Use Aspose.Cells to Show Custom Data Labels from a Cell Range for the First Series of a Column Chart
// Description: Creates a workbook, adds categories, values and custom label texts, inserts a column chart, links the first series to the value range, enables data labels, sets ShowCellRange to true, assigns the label source (C2:C4), applies font styling, and saves the file as an XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart data labels | custom data label range Aspose.Cells | ShowCellRange Aspose.Cells | LinkedSource chart series .NET | column chart custom labels Excel | Aspose.Cells data label styling
// Common Searches: Aspose.Cells link data labels to cell range | How to set custom text for chart series labels in C# | ShowCellRange and LinkedSource example Aspose.Cells | C# chart data labels from cells Aspose | Customize chart label font color Aspose.Cells
// Developer Intent: Generate a column chart where the first series displays data labels taken from a specified worksheet range.
// Use Cases: Display unit‑specific text (e.g., "100 units") on each column bar by referencing cells C2:C4. | Apply distinct font colors or styles to custom data labels for clearer visual emphasis. | Automate Excel reports that combine raw values with formatted label text linked directly to the chart.
// AI Prompts: Write C# code with Aspose.Cells to enable ShowCellRange and set LinkedSource so chart data labels use text from a cell range. | Explain how to change the font color and style of data labels for a chart series in Aspose.Cells for .NET. | Provide step‑by‑step instructions to save an XLSX file after adding a column chart with custom data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelsFromRange
{
    // Creates a workbook, adds categories, values and custom label texts, inserts a column chart, links the first series to the value range, enables data labels, sets ShowCellRange to true, assigns the label source (C2:C4), applies font styling, and saves the file as an XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Custom label texts that will be used as data labels
            sheet.Cells["C2"].PutValue("100 units");
            sheet.Cells["C3"].PutValue("200 units");
            sheet.Cells["C4"].PutValue("300 units");

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values) and categories
            chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

            // Access the first series
            Series firstSeries = chart.NSeries[0];

            // Enable data labels and configure them to use the custom cell range
            firstSeries.DataLabels.ShowValue = true;          // Show the numeric value (optional)
            firstSeries.DataLabels.ShowCellRange = true;     // Use cell range for label text
            firstSeries.DataLabels.LinkedSource = "C2:C4";   // Range containing custom label texts
            firstSeries.DataLabels.Font.Color = Color.Blue; // Example styling

            // Save the workbook
            workbook.Save("ChartWithCustomDataLabels.xlsx", SaveFormat.Xlsx);
        }
    }
}
