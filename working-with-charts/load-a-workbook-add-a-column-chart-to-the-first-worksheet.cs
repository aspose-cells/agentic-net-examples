// Title: C# – Add a Column Chart to the First Worksheet of an Existing Workbook with Aspose.Cells
// Description: Loads an existing Excel file, writes sample data to columns A and B, inserts a column chart on the first sheet (rows 5‑20, columns 0‑5), binds the series to B2:B5 and categories to A2:A5, sets a chart title, and saves the result as a new workbook.
// Keywords: Aspose.Cells column chart C# | add chart to existing workbook .NET | Aspose.Cells NSeries example | set chart title Aspose.Cells | save workbook with chart Aspose | GitHub Aspose.Cells chart sample | C# Excel chart automation
// Common Searches: how to insert a column chart in Aspose.Cells C# | Aspose.Cells create chart from existing workbook | C# code for adding a column chart to first sheet | Aspose.Cells set data range for chart | example of saving workbook with chart Aspose
// Developer Intent: Programmatically add a column chart to the first worksheet of a loaded workbook and write the updated file.
// Use Cases: Generate a sales‑by‑region column chart from a template workbook before sending reports. | Automate visual summaries for monthly KPI data across multiple workbooks. | Create a reusable chart‑insertion routine for Excel dashboards built with Aspose.Cells.
// AI Prompts: Show C# code that uses Aspose.Cells to add a clustered column chart to the first sheet of an existing Excel file, using data from columns A and B. | Provide an example that positions the chart, sets its title, and saves the workbook as output.xlsx. | Explain how to switch the chart type to stacked column and modify its size and location programmatically with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing Excel file, writes sample data to columns A and B, inserts a column chart on the first sheet (rows 5‑20, columns 0‑5), binds the series to B2:B5 and categories to A2:A5, sets a chart title, and saves the result as a new workbook.
class AddColumnChart
{
    static void Main()
    {
        // Load an existing workbook from file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart (if the sheet is empty)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 5; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B5", true);          // Values
        chart.NSeries.CategoryData = "A2:A5";      // Categories

        // Optional: set a title for the chart
        chart.Title.Text = "Sample Column Chart";

        // Save the workbook with the newly added chart
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
