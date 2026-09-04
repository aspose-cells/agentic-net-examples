// Title: Retrieve the parent worksheet of a chart with Aspose.Cells Chart.Worksheet in C#
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, selects a chart (by index or name), and uses the Chart.Worksheet property to output the worksheet's name. | Create a C# example that loads a .xlsx file, accesses the first chart on the first sheet, and prints the name of the sheet that owns the chart via Chart.Worksheet. | Write a C# snippet that finds a chart by its title in an Aspose.Cells workbook and returns the containing worksheet object.
// Common Searches: Aspose.Cells C# get worksheet of a specific chart | How to use Chart.Worksheet to identify chart's sheet in .NET | Find parent worksheet for an Excel chart using Aspose.Cells library | Retrieve chart's sheet name with Aspose.Cells Chart.Worksheet property | C# Aspose.Cells example to locate chart's worksheet by chart name
// Tags: Aspose.Cells Chart.Worksheet property | C# retrieve chart parent worksheet | Excel chart location Aspose.Cells | chart worksheet identification .NET | load workbook access chart worksheet

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an Excel workbook, accesses a chart, uses the Chart.Worksheet property to obtain the worksheet that contains the chart, and prints the worksheet's name.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Verify the input file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Ensure the workbook has at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("The worksheet contains no charts.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Get the worksheet that holds this chart
            Worksheet chartWorksheet = chart.Worksheet;

            // Display the name of the worksheet containing the chart
            Console.WriteLine("The chart is located in worksheet: " + chartWorksheet.Name);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
