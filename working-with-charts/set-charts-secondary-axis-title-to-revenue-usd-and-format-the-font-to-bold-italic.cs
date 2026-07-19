// Title: Aspose.Cells C# – Set secondary Y‑axis title to "Revenue (USD)" with bold italic font
// Description: Shows how to create a column chart in Aspose.Cells for .NET, plot a series on the secondary Y‑axis, make the secondary axis title visible, assign the text "Revenue (USD)", and format the title font as bold and italic before saving the workbook.
// Keywords: Aspose.Cells | C# chart | secondary axis title | bold italic font | dual‑axis column chart | Aspose.Cells .NET | format chart axis | secondary Y axis label
// Common Searches: Aspose.Cells set secondary axis title | C# chart secondary Y axis bold italic | How to format secondary axis label in Aspose.Cells | Add title to secondary axis Aspose.Cells .NET | Dual axis chart formatting Aspose.Cells
// Developer Intent: Add a visible title to a chart’s secondary Y‑axis and style the title with bold and italic font using Aspose.Cells for .NET.
// Use Cases: Financial statements that compare units sold and revenue on separate axes, emphasizing the revenue axis label. | Marketing dashboards displaying traffic and conversion value where the secondary axis title clarifies monetary units. | Operational reports that plot production volume and cost, using a formatted secondary axis title for clear distinction.
// AI Prompts: Generate C# code with Aspose.Cells to set a secondary Y‑axis title to a custom string and apply bold italic styling. | Provide an Aspose.Cells example that changes the secondary axis title color and font size in addition to bold and italic. | Show how to hide or remove the secondary axis title after it has been created in an Aspose.Cells chart. | Write a script that adds a secondary axis title, then aligns it to the right and adds a background fill using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a column chart in Aspose.Cells for .NET, plot a series on the secondary Y‑axis, make the secondary axis title visible, assign the text "Revenue (USD)", and format the title font as bold and italic before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for two series
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");

        worksheet.Cells["B1"].PutValue("Series 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["B4"].PutValue(300);

        worksheet.Cells["C1"].PutValue("Series 2");
        worksheet.Cells["C2"].PutValue(5000);
        worksheet.Cells["C3"].PutValue(3000);
        worksheet.Cells["C4"].PutValue(1000);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Add the two data series and set category data
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary (right) Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Configure the secondary axis title
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Revenue (USD)";
        secondaryAxis.Title.IsVisible = true;
        secondaryAxis.Title.Font.IsBold = true;
        secondaryAxis.Title.Font.IsItalic = true;

        // Save the workbook
        workbook.Save("SecondaryAxisTitle.xlsx");
    }
}
