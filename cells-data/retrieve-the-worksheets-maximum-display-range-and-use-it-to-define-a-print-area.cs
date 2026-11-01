// Title: Set worksheet print area to its maximum display range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet print area to the used range | How to retrieve MaxDisplayRange and assign it as print area in Aspose.Cells | Aspose.Cells .NET example for defining print area based on data range | Set Excel print area programmatically using MaxDisplayRange with Aspose.Cells
// Tags: Aspose.Cells MaxDisplayRange print area | C# worksheet page setup print area | define print area from used cells Aspose.Cells | Excel workbook set print area programmatically | Aspose.Cells retrieve maximum display range

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, gets the first worksheet, obtains its MaxDisplayRange (the area containing data or formatting), converts the range to an A1‑style reference, assigns that reference to the worksheet's PageSetup.PrintArea, and saves the workbook with the updated print settings.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve the maximum display range (the area that contains data or formatting)
                AsposeRange maxRange = sheet.Cells.MaxDisplayRange;

                // Convert the range to an A1‑style reference string (e.g., "A1:D20")
                string printArea = maxRange.RefersTo; // or maxRange.GetRefersTo()

                // Set the print area for the worksheet
                sheet.PageSetup.PrintArea = printArea;

                // Save the workbook with the new print area settings
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
