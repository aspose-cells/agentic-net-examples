// Title: Insert a stacked bar chart into an existing XLSX workbook using Aspose.Cells for .NET and save the result
// AI Prompts: Create a BarStacked chart on the first worksheet of a loaded workbook, using the range A1:B5, set the title to "Stacked Bar Chart", and save the workbook to a new file with Aspose.Cells in C#. | Load an existing Excel file, add a stacked bar chart with data from A1:B5, customize the chart title, and export the modified workbook to a different path using Aspose.Cells for .NET. | Generate a stacked bar chart programmatically in C#, attach it to the first sheet of an XLSX workbook, and write the updated workbook to disk with Aspose.Cells.
// Common Searches: how to add a stacked bar chart to an existing Excel file using Aspose.Cells C# | Aspose.Cells example for inserting BarStacked chart into loaded workbook | C# code to load XLSX, create stacked bar chart from range A1:B5, and save file | save workbook after adding chart Aspose.Cells .NET | chart.NSeries.Add range A1:B5 Aspose.Cells stacked bar example
// Tags: insert BarStacked chart Aspose.Cells | load XLSX workbook C# Aspose | define chart data range A1:B5 Aspose | set chart title Aspose.Cells | export modified workbook .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads 'input.xlsx', adds a stacked bar chart on the first worksheet using data from A1:B5, sets the chart title to 'Stacked Bar Chart', and saves the updated workbook as 'output.xlsx' with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a stacked bar chart (upper‑left row, column, lower‑right row, column)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 10);

            // Retrieve the newly added chart
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart series (example range A1:B5)
            chart.NSeries.Add("A1:B5", true);

            // Set the chart title
            chart.Title.Text = "Stacked Bar Chart";

            // Save the workbook with the new chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
