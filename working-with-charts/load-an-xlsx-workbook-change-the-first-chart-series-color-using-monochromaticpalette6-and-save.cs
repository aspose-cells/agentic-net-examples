// Title: Change the first series color of the first chart in an XLSX workbook with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file using Aspose.Cells, locate the first worksheet's first chart, set the foreground color of its first series to a specific RGB value (e.g., #FF9933), and save the workbook to a new file. | Using the Aspose.Cells .NET API, access the first chart series in a workbook, apply a monochromatic palette shade to the series area, and write the updated workbook back to disk.
// Common Searches: how to change chart series color using Aspose.Cells C# | Aspose.Cells set first series foreground color in Excel chart | modify Excel chart series color programmatically .NET | apply custom RGB to chart series with Aspose.Cells | load workbook edit chart series style save Aspose.Cells
// Tags: modify chart series color Aspose.Cells | set first series foreground color C# | apply monochromatic palette to Excel chart series | load and save XLSX workbook Aspose.Cells | first chart series styling .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads 'input.xlsx', verifies the presence of at least one chart and series, changes the foreground color of the first series to an orange RGB shade, and saves the modified workbook as 'output.xlsx', handling missing files or chart elements gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure the chart contains at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("No series found in the chart.");
                return;
            }

            // Access the first series
            Series series = chart.NSeries[0];

            // Set the series color (using a sample monochrome-like color)
            series.Area.ForegroundColor = Color.FromArgb(0xFF, 0x99, 0x33);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
