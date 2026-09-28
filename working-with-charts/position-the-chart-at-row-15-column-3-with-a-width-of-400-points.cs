// Title: How to position a column chart at cell C15 and set its width to 400 points using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, moves it so its upper‑left corner aligns with cell C15, and sets the chart object's width to 400 points. | Show how to relocate an existing Aspose.Cells chart to row 15, column 3 and change its width property to 400 points in an .xlsx workbook.
// Common Searches: Aspose.Cells C# move chart to cell C15 | set chart width in points using Aspose.Cells .NET | position Excel chart at specific row and column with Aspose.Cells | Aspose.Cells chart.Move method example | adjust column chart size programmatically in C#
// Tags: Aspose.Cells chart.Move positioning | Aspose.Cells chart width property | C# Aspose.Cells set chart size | Aspose.Cells column chart placement | Aspose.Cells workbook chart positioning

using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds sample data, inserts a column chart, moves the chart so its upper‑left corner starts at row 15 column 3 (cell C15), sets the chart width to 400 points, and saves the file as ChartPositioned.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);

        // Add a chart with an initial position
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Move the chart so its upper‑left corner is at row 15, column 3
        // Bottom row and right column are set to keep a reasonable size
        chart.Move(15, 3, 25, 10);

        // Set the chart width to 400 points
        chart.ChartObject.Width = 400;

        // Save the workbook
        workbook.Save("ChartPositioned.xlsx", SaveFormat.Xlsx);
    }
}
