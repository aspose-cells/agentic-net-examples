// Title: Add a stacked bar chart to an existing XLSX workbook using Aspose.Cells for .NET and save the result
// AI Prompts: Write C# code that opens a specified XLSX file with Aspose.Cells, inserts a BarStacked chart on the first worksheet using a custom data range, sets a chart title, and saves the workbook to a new file. | Adjust the provided Aspose.Cells example to place the stacked bar chart between rows 2‑10 and columns 1‑4 and to use the range C2:D6 as the data source. | Describe how to update the title and series of a stacked bar chart after loading an existing workbook with Aspose.Cells.
// Common Searches: how to add a stacked bar chart to an existing Excel file using Aspose.Cells C# | Aspose.Cells .NET insert chart into workbook while keeping original data | C# example creating BarStacked chart from range A1:B5 with Aspose.Cells | save workbook after adding chart Aspose.Cells .NET | change chart position in Aspose.Cells after loading an existing workbook
// Tags: Aspose.Cells stacked bar chart insertion | Aspose.Cells workbook loading | Aspose.Cells chart series definition | Aspose.Cells chart layout configuration | Aspose.Cells workbook export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, adds a BarStacked chart on the first worksheet using range A1:B5, sets the chart title, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a stacked bar chart (upper‑left row, column, lower‑right row, column)
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data range for the chart (example: A1:B5)
            chart.NSeries.Add("A1:B5", true);

            // Set chart title
            chart.Title.Text = "Stacked Bar Chart";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
