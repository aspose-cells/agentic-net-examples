// Title: How to set a custom caption for an existing slicer in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to change the Caption property of the first worksheet slicer to a custom text. | Programmatically assign a new label and optionally a new name to a slicer before saving the workbook with Aspose.Cells.
// Common Searches: asp.net change slicer caption with Aspose.Cells | c# update slicer label in existing .xlsx using Aspose.Cells | modify slicer caption and name in Excel workbook via Aspose.Cells .NET
// Tags: Aspose.Cells slicer caption property | C# update Excel slicer label | Aspose.Cells modify slicer name | programmatic Excel slicer customization .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerCaptionExample
{
    // The example loads an existing Excel file, checks the first worksheet for slicers, sets the Caption of the first slicer to a custom string (optionally changes its Name), and saves the workbook, demonstrating how to customize slicer captions with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

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

                // Check if the worksheet contains any slicers
                if (worksheet.Slicers.Count > 0)
                {
                    // Retrieve the first slicer
                    Slicer slicer = worksheet.Slicers[0];

                    // Set a custom caption for the slicer
                    slicer.Caption = "Custom Slicer Caption";

                    // Optionally set the slicer's name
                    // slicer.Name = "MySlicer";
                }
                else
                {
                    Console.WriteLine("No slicers found in the worksheet.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
