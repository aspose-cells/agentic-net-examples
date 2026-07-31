// Title: Add a Stacked Bar Chart to an Existing XLSX Workbook with Aspose.Cells for .NET
// Description: Loads an existing XLSX file, inserts a BarStacked chart into the first worksheet (rows 5‑25, columns 1‑10), defines the data series as A2:D5, and saves the workbook as a new file using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | stacked bar chart | ChartType.BarStacked | add chart to workbook | load Excel file | save workbook with chart | .NET Excel automation | NSeries range syntax | Excel chart example
// Common Searches: Aspose.Cells add stacked bar chart C# | how to insert BarStacked chart into existing Excel file .NET | Aspose.Cells NSeries range syntax for charts | save workbook after adding chart Aspose.Cells | C# code example for positioning charts in Aspose.Cells
// Developer Intent: Insert a stacked bar chart into an existing Excel workbook and persist the changes.
// Use Cases: Generate a sales‑by‑region report by loading a template workbook, adding a stacked bar chart that visualizes A2:D5, and exporting the result. | Automate monthly KPI dashboards by programmatically adding a BarStacked chart to a workbook that already contains raw metrics, then distributing the updated file. | Create a reusable chart template where the stacked bar chart is added once; subsequent runs only refresh the data range before saving.
// AI Prompts: Write C# code with Aspose.Cells that adds a BarStacked chart to an existing workbook, using a dynamic data range and custom chart position. | Explain how to configure NSeries source and category labels for a stacked bar chart in Aspose.Cells. | Show how to modify the size, colors, and axis titles of a stacked bar chart after it has been added to a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, inserts a BarStacked chart into the first worksheet (rows 5‑25, columns 1‑10), defines the data series as A2:D5, and saves the workbook as a new file using Aspose.Cells in C#.
class AddStackedBarChart
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Load the workbook from the file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a stacked bar chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn (zero‑based indices)
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 1, 25, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        // Adjust the range to match the actual data in the worksheet
        chart.NSeries.Add("=Sheet1!$A$2:$D$5", true);

        // Save the modified workbook to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
