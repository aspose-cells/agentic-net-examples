// Title: C# – Set PivotChart Legend to Bottom in an Existing XLSX Workbook with Aspose.Cells
// Description: Loads an existing XLSX file, verifies the first chart is linked to a PivotTable, changes its legend position to the bottom using Aspose.Cells for .NET, and saves the workbook to a new file.
// Keywords: Aspose.Cells C# PivotChart legend position | set chart legend bottom Aspose.Cells | modify PivotChart legend .NET | Aspose.Cells change chart legend location | C# update PivotChart legend | Aspose.Cells chart formatting
// Common Searches: Aspose.Cells set PivotChart legend to bottom | C# change legend position of a PivotChart | How to move PivotChart legend in Aspose.Cells | Update chart legend location in existing XLSX with Aspose | Aspose.Cells example: legend position bottom
// Developer Intent: Programmatically move the legend of a PivotChart to the bottom of the chart area and persist the change in the workbook.
// Use Cases: Standardize legend placement across generated reports for clearer dashboards. | Automate post‑processing of workbooks that contain PivotCharts to meet corporate style guides. | Ensure a chart is linked to a PivotTable before applying formatting to avoid runtime errors.
// AI Prompts: Generate C# code using Aspose.Cells that opens an XLSX file, finds the first PivotChart, sets its legend position to Bottom, and saves the workbook. | Provide an Aspose.Cells snippet that checks a chart's PivotSource before changing the legend position, with proper exception handling. | Create a reusable method that accepts input and output paths and updates the legend position of all PivotCharts in a workbook to the bottom.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, verifies the first chart is linked to a PivotTable, changes its legend position to the bottom using Aspose.Cells for .NET, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The input file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Assume the workbook already contains a PivotChart.
            // Retrieve the first chart in the worksheet.
            if (worksheet.Charts.Count > 0)
            {
                try
                {
                    Chart chart = worksheet.Charts[0];

                    // Verify that the chart is linked to a PivotTable.
                    if (!string.IsNullOrEmpty(chart.PivotSource))
                    {
                        // Set the legend position to the bottom of the chart.
                        // Use LegendPositionType enum for compatibility with various Aspose.Cells versions.
                        chart.Legend.Position = LegendPositionType.Bottom;
                    }
                }
                catch (Exception exChart)
                {
                    Console.WriteLine($"Chart processing error: {exChart.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
