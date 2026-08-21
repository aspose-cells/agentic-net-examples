// Title: Add a multi‑series line chart for quarterly sales with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, writes quarter labels and sales figures for two products, inserts a line chart spanning rows 6‑20 and columns A‑J, defines series (B2:C5) and categories (A2:A5), sets a chart title, and saves the file as QuarterlyLineChart.xlsx.
// Keywords: Aspose.Cells | C# | line chart | quarterly sales | multiple series | chart series range | category axis | save as xlsx | Excel chart automation
// Common Searches: asp.net add line chart Aspose.Cells | asp.net line chart multiple series | asp.net set category axis Aspose.Cells | asp.net create quarterly sales chart | asp.net export chart to pdf Aspose.Cells
// Developer Intent: Insert a line chart that visualizes Product A and Product B sales across four quarters on the same worksheet.
// Use Cases: Build a sales performance report that shows quarterly trends for several products. | Automate financial dashboards by programmatically generating line charts from worksheet data. | Create reusable chart templates that pull category and series values directly from cells.
// AI Prompts: Generate C# code to add a line chart with two series from columns B and C and set the category axis to column A using Aspose.Cells. | Show how to customize line colors, markers, and legend entries for a chart created with Aspose.Cells for .NET. | Explain how to export a workbook that contains a line chart to PDF using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, writes quarter labels and sales figures for two products, inserts a line chart spanning rows 6‑20 and columns A‑J, defines series (B2:C5) and categories (A2:A5), sets a chart title, and saves the file as QuarterlyLineChart.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Quarter");
        sheet.Cells["B1"].PutValue("Product A");
        sheet.Cells["C1"].PutValue("Product B");

        // Sample data for four quarters
        string[] quarters = { "Q1", "Q2", "Q3", "Q4" };
        double[] salesA = { 12000, 15000, 13000, 17000 };
        double[] salesB = { 10000, 14000, 11000, 16000 };

        // Fill the worksheet with the data
        for (int i = 0; i < quarters.Length; i++)
        {
            int row = i + 2; // Data starts at row 2 (index 1)
            sheet.Cells[row, 0].PutValue(quarters[i]); // Column A
            sheet.Cells[row, 1].PutValue(salesA[i]);   // Column B
            sheet.Cells[row, 2].PutValue(salesB[i]);   // Column C
        }

        // Add a line chart to the worksheet (rows 6‑20, columns 1‑10)
        int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 9);
        Chart chart = sheet.Charts[chartIndex];

        // Define the series data (Product A and Product B) – series are taken by column
        chart.NSeries.Add("B2:C5", true);
        // Define the category axis data (quarters)
        chart.NSeries.CategoryData = "A2:A5";

        // Set a descriptive title
        chart.Title.Text = "Quarterly Sales Comparison";

        // Save the workbook with the chart
        workbook.Save("QuarterlyLineChart.xlsx", SaveFormat.Xlsx);
    }
}
