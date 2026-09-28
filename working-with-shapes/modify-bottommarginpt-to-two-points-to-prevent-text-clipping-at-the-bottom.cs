// Title: Set worksheet bottom margin to 2 points with Aspose.Cells for .NET to avoid text clipping
// AI Prompts: Generate C# code that opens an existing Excel file using Aspose.Cells, sets the worksheet PageSetup.BottomMargin to 2 points, and saves the file. | Show how to change the bottom margin of a worksheet to two points in Aspose.Cells to prevent content from being cut off. | Provide a minimal .NET example that adjusts the BottomMargin property of a worksheet before exporting to PDF.
// Common Searches: Aspose.Cells C# set bottom margin to 2 points to stop clipping | how to change worksheet bottom margin in .NET Excel library | prevent bottom text cut off when printing Excel with Aspose.Cells | adjust page setup margins programmatically using Aspose.Cells | example code for modifying BottomMargin property in Aspose.Cells workbook
// Tags: Aspose.Cells set bottom margin points | C# Excel worksheet page setup margins | adjust worksheet bottom margin Aspose.Cells | prevent bottom text clipping Excel | modify page layout margins .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, sets the first worksheet's PageSetup.BottomMargin to 2 points to prevent bottom text clipping, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (or any specific worksheet you need)
                Worksheet sheet = workbook.Worksheets[0];

                // Set the bottom margin to 2 points to avoid clipping
                sheet.PageSetup.BottomMargin = 2;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
