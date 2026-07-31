// Title: Set Custom Number Formats on PivotChart Axes with Aspose.Cells for .NET
// Description: Loads an XLSX workbook, checks that the first worksheet has a chart, assigns a currency format to the value‑axis and a date format to the category‑axis of a PivotChart, and saves the modified file.
// Keywords: Aspose.Cells PivotChart axis formatting | C# custom number format chart | Excel value axis currency format | category axis date format Aspose | modify chart tick labels .NET | save workbook after chart changes | Aspose.Cells chart number format
// Common Searches: how to format pivotchart value axis with Aspose.Cells | set date format on chart category axis C# | apply currency format to Excel chart axis using Aspose | Aspose.Cells change chart tick label format | save workbook after updating chart formatting
// Developer Intent: Apply specific number formats to a PivotChart’s axes and write the updated workbook to disk.
// Use Cases: Show sales amounts as $#,##0.00 on the value axis of a quarterly revenue PivotChart. | Display month‑day labels (mmm dd) on the category axis of a timeline chart. | Validate worksheet and chart existence before formatting to avoid runtime errors.
// AI Prompts: Generate C# code with Aspose.Cells that sets a currency format on a PivotChart value axis and saves the workbook. | Create error‑handling logic that confirms a worksheet contains at least one chart before applying axis formats. | Provide an example that formats the category axis of a PivotChart with a month‑day pattern and exports the file as XLSX.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, checks that the first worksheet has a chart, assigns a currency format to the value‑axis and a date format to the category‑axis of a PivotChart, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook contains at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("Error: The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("Error: No charts were found on the first worksheet.");
                return;
            }

            // Retrieve the first chart – assumed to be the PivotChart
            Chart pivotChart = worksheet.Charts[0];

            // Apply a custom number format to the value‑axis tick labels (e.g., currency)
            pivotChart.ValueAxis.TickLabels.NumberFormat = "$#,##0.00";

            // Optional: Apply a custom number format to the category‑axis tick labels (e.g., month and day)
            pivotChart.CategoryAxis.TickLabels.NumberFormat = "mmm dd";

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
