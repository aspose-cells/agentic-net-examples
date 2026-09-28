// Title: Create and bind a PivotChart to an existing XLS workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xls file with Aspose.Cells, checks for a PivotTable, adds a column PivotChart linked to that PivotTable, sets a custom title, and saves the workbook. | Generate a .NET snippet that loads a legacy Excel file, creates a chart range (rows 10‑20, columns A‑H), binds the chart to the first PivotTable, and exports the result to a new .xls file. | Provide a C# example that validates the presence of a PivotTable, adds a Chart object of type Column, connects its series to the PivotTable name, and persists the changes using Aspose.Cells.
// Common Searches: Aspose.Cells C# add PivotChart to existing .xls workbook | How to bind a chart to a PivotTable in a legacy Excel file using Aspose.Cells | Create a column chart from the first PivotTable in an .xls file with Aspose.Cells .NET | C# example for verifying PivotTable before adding chart with Aspose.Cells | Save modified workbook after adding PivotChart using Aspose.Cells
// Tags: create column pivotchart in legacy xls via Aspose.Cells | associate chart series with pivot table using Aspose.Cells | load existing xls workbook and add chart C# | Aspose.Cells example for chart‑pivot binding | modify worksheet to include pivotchart .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// The C# program loads 'input.xls', ensures at least one PivotTable exists, adds a column PivotChart positioned between rows 10‑20 and columns A‑H, binds the chart series to the first PivotTable, sets a visible title, and saves the updated workbook as 'output.xls', with error handling for missing files and exceptions.
class PivotChartExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a PivotTable exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTable found in the worksheet.");
                return;
            }

            // Retrieve the first PivotTable
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a new chart (type, upper-left row/col, lower-right row/col)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 0, 20, 7);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind the chart to the PivotTable
            chart.NSeries.Add(pivotTable.Name, true);

            // Set chart title
            chart.Title.Text = "PivotChart based on " + pivotTable.Name;
            chart.Title.IsVisible = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
