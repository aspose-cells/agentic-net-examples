// Title: Change the first chart series range to B2:B6 and save the workbook as ChartReport.xlsx using Aspose.Cells for .NET (C#)
// AI Prompts: Load Input.xlsx, locate the first chart on the first worksheet, set its first series values to the range B2:B6, and save the workbook as ChartReport.xlsx with Aspose.Cells in C#. | Create C# code that checks whether an Excel file exists, opens it with Aspose.Cells, updates the first chart's series data source to B2:B6, and writes the modified workbook to a new file named ChartReport.xlsx.
// Common Searches: asp.net change chart series data range to B2:B6 using Aspose.Cells | save modified workbook to a new Excel file with Aspose.Cells C# | how to update first chart on a worksheet in an existing workbook Aspose.Cells | verify Excel file existence before loading with Aspose.Cells in C#
// Tags: update chart series values Aspose.Cells | save workbook to new Excel file C# | check input file existence Aspose.Cells | modify first worksheet chart Aspose.Cells | set NSeries range B2:B6

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program verifies that Input.xlsx exists, opens it with Aspose.Cells, accesses the first worksheet, updates the first chart's first series to use the data range B2:B6, and saves the modified workbook as ChartReport.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "ChartReport.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count > 0)
            {
                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Modify the first series' values if a series exists
                if (chart.NSeries.Count > 0)
                {
                    // Set new data range for the series (e.g., B2:B6)
                    chart.NSeries[0].Values = "B2:B6";
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
