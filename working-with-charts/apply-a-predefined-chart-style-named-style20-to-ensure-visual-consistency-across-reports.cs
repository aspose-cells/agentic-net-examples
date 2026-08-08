// Title: Apply Predefined Chart Style Style20 to a Column Chart with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, sets its series and categories, applies the built‑in Style20 (chart.Style = 20) and saves the file as ChartWithStyle20.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | chart style | Style20 | predefined chart style | chart.Style property | column chart | Excel automation
// Common Searches: Aspose.Cells set chart style | How to use Style20 in Aspose.Cells | Apply predefined chart style C# Aspose | ChartStyle enumeration Aspose.Cells | Change chart appearance Aspose.Cells .NET
// Developer Intent: Programmatically apply the built‑in Style20 to a chart to achieve consistent visual formatting across reports.
// Use Cases: Generate financial dashboards where every column chart follows the corporate Style20 theme. | Batch‑process existing workbooks to update all charts to Style20 for uniform branding. | Create template workbooks that automatically apply Style20 to new charts added by end‑users.
// AI Prompts: Show how to apply a different built‑in chart style, such as Style15, to a line chart using Aspose.Cells for .NET. | Explain how to retrieve the full list of available chart style IDs in Aspose.Cells and select one at runtime. | Provide C# code that loops through all charts in a workbook and sets each chart's Style property to 20.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, sets its series and categories, applies the built‑in Style20 (chart.Style = 20) and saves the file as ChartWithStyle20.xlsx using Aspose.Cells for .NET.
class ApplyChartStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", false);
        chart.NSeries.CategoryData = "A2:A4";

        // Apply the predefined chart style named Style20 (value 20)
        chart.Style = 20;

        // Save the workbook with the styled chart
        workbook.Save("ChartWithStyle20.xlsx");
    }
}
