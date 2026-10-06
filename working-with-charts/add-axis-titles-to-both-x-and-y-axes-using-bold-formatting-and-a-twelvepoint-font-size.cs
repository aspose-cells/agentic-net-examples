// Title: How to apply bold 12‑point titles to X‑ and Y‑axes of a column chart with Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart in a new workbook and set the X‑axis title to "Months" and the Y‑axis title to "Sales" using bold 12‑point font with Aspose.Cells in C#. | Use Aspose.Cells to format chart axis titles by enabling bold style and specifying a 12‑point font size for both category and value axes. | Generate an Excel file that contains a column chart whose axis titles are programmatically styled with bold 12‑point text via the Aspose.Cells API.
// Common Searches: C# Aspose.Cells set 12‑point bold text for chart X and Y axis titles | How to add and format axis titles in a column chart using Aspose.Cells for .NET | Aspose.Cells chart axis title styling example in C# | Programmatically set category and value axis titles with font size in Aspose.Cells workbook
// Tags: Aspose.Cells chart axis title font styling | C# column chart axis title bold | Aspose.Cells set axis title size 12 | Excel workbook chart axis title formatting | Aspose.Cells add X and Y axis titles

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart, assigns category and value ranges, sets the X‑axis title to "Months" and the Y‑axis title to "Sales" with bold 12‑point fonts, and saves the file as ChartWithAxisTitles.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Configure X‑axis (category axis) title
        chart.CategoryAxis.Title.Text = "Months";
        chart.CategoryAxis.Title.Font.IsBold = true;
        chart.CategoryAxis.Title.Font.Size = 12;

        // Configure Y‑axis (value axis) title
        chart.ValueAxis.Title.Text = "Sales";
        chart.ValueAxis.Title.Font.IsBold = true;
        chart.ValueAxis.Title.Font.Size = 12;

        // Save the workbook to a file
        workbook.Save("ChartWithAxisTitles.xlsx");
    }
}
