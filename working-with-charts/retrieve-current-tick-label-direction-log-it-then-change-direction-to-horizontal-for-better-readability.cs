// Title: Aspose.Cells C# – Get and Set Chart Category Axis Tick Labels Direction to Horizontal
// Description: Creates a workbook with a column chart, reads the current CategoryAxis.TickLabels.DirectionType, logs it, then changes the direction to Horizontal and saves the file.
// Keywords: Aspose.Cells | C# | Chart | CategoryAxis | TickLabels | DirectionType | ChartTextDirectionType | horizontal tick labels | read tick label orientation | modify chart axis labels | Excel automation
// Common Searches: Aspose.Cells get tick label direction | set chart axis tick labels horizontal Aspose.Cells | ChartTextDirectionType example C# | read and change chart tick label orientation | Aspose.Cells chart axis label direction
// Developer Intent: Read the current tick label orientation of a chart’s category axis and change it to horizontal for better readability.
// Use Cases: Log original tick label direction during automated report generation. | Standardize axis label orientation across Excel files produced by a service. | Provide a helper method that forces horizontal labels based on chart size or language settings. | Debug rendering issues by inspecting the label direction before modification.
// AI Prompts: Write C# code using Aspose.Cells to output the current Chart.CategoryAxis.TickLabels.DirectionType and then set it to ChartTextDirectionType.Horizontal. | Explain how ChartTextDirectionType affects axis label rendering and demonstrate toggling between vertical, rotated, and horizontal orientations. | Generate a reusable function that accepts a Chart object, returns its previous TickLabels.DirectionType, and forces the direction to Horizontal.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook with a column chart, reads the current CategoryAxis.TickLabels.DirectionType, logs it, then changes the direction to Horizontal and saves the file.
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

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);               // Values
        chart.NSeries.CategoryData = "A2:A4";           // Categories

        // Access the tick labels of the category (X) axis
        TickLabels tickLabels = chart.CategoryAxis.TickLabels;

        // Retrieve and log the current direction of the tick labels
        ChartTextDirectionType currentDirection = tickLabels.DirectionType;
        Console.WriteLine($"Current TickLabels DirectionType: {currentDirection}");

        // Change the direction to Horizontal for better readability
        tickLabels.DirectionType = ChartTextDirectionType.Horizontal;

        // Save the workbook
        workbook.Save("TickLabelsDirectionChanged.xlsx");
    }
}
