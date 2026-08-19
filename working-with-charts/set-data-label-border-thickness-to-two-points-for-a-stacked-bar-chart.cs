// Title: C# Example: Set 2‑pt Data Label Border on a Stacked Bar Chart with Aspose.Cells for .NET
// Description: A concise Aspose.Cells for .NET sample that creates a workbook, fills it with data, adds a stacked bar chart, turns on data labels, makes the label border visible, and sets the border thickness to 2 points before saving the file.
// Keywords: Aspose.Cells C# chart example | stacked bar chart data label border | set border thickness 2 points | .NET Excel automation | chart styling Aspose.Cells | data label border visibility | Excel workbook generation | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells set data label border thickness C# | how to change data label border weight in stacked bar chart Aspose.Cells | C# example for chart data label border visibility | Aspose.Cells .NET chart styling tutorial | set 2 pt border on chart labels Aspose
// Developer Intent: Apply a 2‑point visible border to data labels of each series in a stacked bar chart.
// Use Cases: Enhance label readability in automated Excel reports. | Standardize chart appearance across multiple generated workbooks. | Create consistent visual styling for dashboards that use stacked bar charts.
// AI Prompts: Write C# code with Aspose.Cells to give data labels a 3‑pt red border on a clustered column chart. | Explain how to toggle data label border visibility and adjust weight for all series in an Aspose.Cells chart. | Provide step‑by‑step instructions to customize data label borders for different chart types using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A concise Aspose.Cells for .NET sample that creates a workbook, fills it with data, adds a stacked bar chart, turns on data labels, makes the label border visible, and sets the border thickness to 2 points before saving the file.
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

            // Add a stacked bar chart (BarStacked is the correct enum value)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart
            chart.NSeries.Add("B2:C4", true);          // values
            chart.NSeries.CategoryData = "A2:A4";      // categories

            // Configure data labels for each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                DataLabels labels = chart.NSeries[i].DataLabels;
                labels.ShowValue = true;               // make labels visible
                labels.Border.WeightPt = 2.0;           // set border thickness to 2 points
                labels.Border.IsVisible = true;        // ensure the border is drawn
            }

            // Save the workbook
            workbook.Save("StackedBarDataLabelBorder.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
