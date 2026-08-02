// Title: C# – Insert a Multi‑Series Line Chart for Quarterly Sales Using Aspose.Cells
// Description: Creates a new workbook, fills A1:D5 with quarter labels and sales data for three products, adds a Line chart (rows 7‑25, cols 1‑10), binds series to A2:D5 with headers, sets quarters as the X‑axis, adds a title, and saves as QuarterlySalesLineChart.xlsx.
// Keywords: Aspose.Cells C# line chart | multi series chart .NET | quarterly sales Excel chart | add chart to worksheet Aspose | Aspose.Cells sample code | GitHub Aspose.Cells example | Excel chart automation C#
// Common Searches: Aspose.Cells add line chart C# | multi series line chart example .NET | set category axis to quarters Aspose.Cells | create quarterly sales chart with Aspose.Cells | Aspose.Cells chart data range with headers
// Developer Intent: Generate a line chart that compares sales figures of several products across quarters on a single worksheet using Aspose.Cells for .NET.
// Use Cases: Automated quarterly sales reports with visual trends | Financial dashboards that plot product performance over time | Exporting sales trend charts to Excel for client presentations | Embedding line charts in .NET applications that produce Excel workbooks
// AI Prompts: Add data markers and a legend to the line chart in the Aspose.Cells example. | Show how to export the generated chart as a PNG image instead of embedding it. | Explain configuring a secondary Y‑axis for one series in an Aspose.Cells line chart.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, fills A1:D5 with quarter labels and sales data for three products, adds a Line chart (rows 7‑25, cols 1‑10), binds series to A2:D5 with headers, sets quarters as the X‑axis, adds a title, and saves as QuarterlySalesLineChart.xlsx.
class InsertLineChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample sales data for four quarters
        // Header row
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");
        sheet.Cells["D1"].PutValue("Product C");

        // Quarter labels
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["A5"].PutValue("Q4");

        // Sales figures for Product A
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(170);
        sheet.Cells["B5"].PutValue(200);

        // Sales figures for Product B
        sheet.Cells["C2"].PutValue(80);
        sheet.Cells["C3"].PutValue(130);
        sheet.Cells["C4"].PutValue(160);
        sheet.Cells["C5"].PutValue(190);

        // Sales figures for Product C
        sheet.Cells["D2"].PutValue(100);
        sheet.Cells["D3"].PutValue(140);
        sheet.Cells["D4"].PutValue(180);
        sheet.Cells["D5"].PutValue(210);

        // Add a line chart to the worksheet (rows 7‑25, columns 1‑10)
        int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 1, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.NSeries.Add("A2:D5", true);
        // Set the category (X‑axis) data to the quarter labels
        chart.NSeries.CategoryData = "A2:A5";

        // Optional: set a chart title
        chart.Title.Text = "Quarterly Sales Comparison";

        // Save the workbook with the chart
        workbook.Save("QuarterlySalesLineChart.xlsx");
    }
}
