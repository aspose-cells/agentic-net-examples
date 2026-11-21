// Title: Set a bold chart title "Quarterly Revenue" in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart title text and make it bold | change Excel chart title to Quarterly Revenue using Aspose.Cells | apply bold formatting to chart title in Aspose.Cells .NET | update first chart title in existing workbook with Aspose.Cells C# | programmatically modify chart title font style Aspose.Cells
// Tags: chart title text assignment Aspose.Cells C# | bold font application to Excel chart title | access first chart worksheet Aspose.Cells | Aspose.Cells chart title property configuration | Excel chart title styling with .NET library

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart class

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, retrieves the first chart on the first worksheet, sets its title to "Quarterly Revenue", applies bold formatting to the title font, and saves the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
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

                // Set chart title and apply bold formatting
                chart.Title.Text = "Quarterly Revenue";
                chart.Title.Font.IsBold = true;

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
