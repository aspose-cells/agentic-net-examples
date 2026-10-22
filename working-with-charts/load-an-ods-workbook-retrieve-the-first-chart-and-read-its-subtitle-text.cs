// Title: Read the first chart's subtitle (title) text from an ODS workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an ODS spreadsheet with Aspose.Cells, locates the first chart, and prints its subtitle text (using the Title property when Subtitle is unavailable). | Generate a method that safely loads an ODS file, checks for the presence of charts, and returns the subtitle string of the first chart, with handling for missing files and empty chart collections. | Create a .NET console application that reads an OpenDocument Spreadsheet, extracts the title of the first chart as a subtitle, and outputs it, including proper exception handling.
// Common Searches: how to extract chart subtitle from an ODS file using Aspose.Cells in C# | Aspose.Cells get first chart title text from OpenDocument Spreadsheet | C# read chart caption in ODS workbook with Aspose.Cells | retrieve chart text when Subtitle property is not exposed in Aspose.Cells | load ODS workbook and list chart titles using Aspose.Cells .NET
// Tags: extract chart title ODS Aspose.Cells | chart subtitle extraction C# Aspose.Cells | load OpenDocument Spreadsheet .NET Aspose.Cells | verify chart collection existence Aspose.Cells | handle missing ODS file Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an ODS workbook with Aspose.Cells, verifies the file exists, checks for charts, retrieves the first chart's Title text (used as subtitle), and writes it to the console, with error handling for missing files and empty chart collections.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count > 0)
            {
                try
                {
                    // Retrieve the first chart
                    Chart chart = sheet.Charts[0];

                    // Aspose.Cells does not expose a separate Subtitle property.
                    // Use the chart's Title text as a representative value.
                    string titleText = chart.Title != null ? chart.Title.Text : string.Empty;

                    // Output the title (used here in place of subtitle)
                    Console.WriteLine("Chart title (used as subtitle): " + titleText);
                }
                catch (Exception exChart)
                {
                    Console.WriteLine($"Error processing chart: {exChart.Message}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
