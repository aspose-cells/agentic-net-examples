// Title: C# – Load an XLSX workbook with Aspose.Cells and get the first chart from the first worksheet
// Description: Opens an existing XLSX file, verifies the file and worksheet, checks for charts, retrieves the first Chart object, prints its Name and Type, and optionally saves the workbook. Includes error handling for missing files or absent charts.
// Keywords: Aspose.Cells | C# chart extraction | load workbook | first chart | worksheet charts | retrieve chart object | read chart name | chart type | Aspose.Cells example | XLSX chart C#
// Common Searches: Aspose.Cells get first chart C# | read chart name from XLSX using Aspose.Cells | C# load workbook and list charts | how to check if worksheet has charts Aspose.Cells | sample code for chart extraction Aspose.Cells .NET
// Developer Intent: Obtain the first Chart object from the first worksheet of an XLSX workbook using Aspose.Cells in C#.
// Use Cases: Log chart metadata for audit or reporting. | Validate that a worksheet contains at least one chart before further processing. | Extract chart information to drive dynamic report generation. | Modify properties of the first chart after loading the workbook. | Automate workbook validation in CI/CD pipelines.
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, checks for charts, returns the first chart's Name and Type, and handles missing file or no charts. | Create a robust Aspose.Cells snippet that iterates all charts in a worksheet, prints each chart's Name and Type, and includes graceful error handling. | Show how to copy the first chart to a new worksheet after loading an XLSX file using Aspose.Cells, preserving its formatting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Opens an existing XLSX file, verifies the file and worksheet, checks for charts, retrieves the first Chart object, prints its Name and Type, and optionally saves the workbook. Includes error handling for missing files or absent charts.
class Program
{
    static void Main()
    {
        try
        {
            string inputFile = "input.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            Worksheet worksheet = workbook.Worksheets[0];

            // Check for charts in the first worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
            }
            else
            {
                // Retrieve and display details of the first chart
                Chart firstChart = worksheet.Charts[0];
                Console.WriteLine($"Chart Name: {firstChart.Name}");
                Console.WriteLine($"Chart Type: {firstChart.Type}");
            }

            // Save the workbook (optional)
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
