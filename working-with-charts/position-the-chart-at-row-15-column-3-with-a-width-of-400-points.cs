// Title: Aspose.Cells for .NET – Place a Column Chart at Row 15, Column C and Set Width to 400 pt (C#)
// Description: Shows how to create a workbook, add sample data, insert a column chart, move its upper‑left corner to row 15, column 3 (C), set the chart width to 400 points with Aspose.Cells for .NET, and save the file as ChartPositioned.xlsx.
// Keywords: Aspose.Cells C# chart positioning | Chart.Move Aspose.Cells | set chart width points | Excel chart placement .NET | column chart row 15 column C | Aspose.Cells ChartObject.WidthPt | position chart by cell | Aspose.Cells example
// Common Searches: Aspose.Cells move chart to specific cell C# | How to set chart width in points using Aspose.Cells | Chart.Move parameters Aspose.Cells .NET | Place Excel chart at row 15 column 3 with Aspose | C# code for chart positioning in Aspose.Cells
// Developer Intent: Position a chart at a precise worksheet cell and define its width in points.
// Use Cases: Standardize layout of sales dashboards by anchoring column charts to row 15, column C with a fixed 400 pt width. | Automate monthly financial reports where charts must align to a predefined grid for consistent branding. | Create reusable Excel templates that insert charts at exact coordinates for multi‑regional deployments.
// AI Prompts: Generate C# code using Aspose.Cells to move an existing chart to row 15, column 3 and set its width to 400 points. | Show how to position a line chart at row 20, column 5 with a height of 300 points using Aspose.Cells for .NET. | Explain the mapping of Chart.Move arguments to worksheet rows and columns and how to convert point measurements to pixel dimensions.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add sample data, insert a column chart, move its upper‑left corner to row 15, column 3 (C), set the chart width to 400 points with Aspose.Cells for .NET, and save the file as ChartPositioned.xlsx.
class Program
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
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart (initial position is arbitrary)
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Move the chart so its upper‑left corner is at row 15, column 3.
        // Bottom row and right column are chosen to give the chart enough height.
        chart.Move(15, 3, 25, 10);

        // Set the chart width to 400 points.
        chart.ChartObject.WidthPt = 400;

        // Save the workbook
        workbook.Save("ChartPositioned.xlsx");
    }
}
