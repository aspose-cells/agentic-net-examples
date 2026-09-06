// Title: Apply a solid green fill to a range of progress cells in an existing Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an existing workbook, creates a style with a solid green background, and applies the fill only to cells A2:A20 using Aspose.Cells StyleFlag to preserve other formatting. | Show how to use Aspose.Cells to verify a file's existence, define a cell range, and apply cell shading without altering fonts or borders in a .NET application.
// Common Searches: Aspose.Cells C# set solid background color for a specific cell range | how to apply only cell shading with StyleFlag in Aspose.Cells | C# example to color progress bar cells in Excel using Aspose.Cells | apply green fill to range A2:A20 without changing fonts Aspose.Cells
// Tags: Aspose.Cells solid fill for cell range | C# StyleFlag cell shading Aspose.Cells | progress cells background color .NET | apply green background to Excel range using Aspose.Cells | cell range formatting without affecting fonts Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The program checks for the input workbook, loads it, defines the range A2:A20 on the first worksheet, creates a style with a solid green fill, applies only the cell shading to that range using a StyleFlag, and saves the updated workbook.
    class ProgressCellGradient
    {
        static void Main()
        {
            try
            {
                const string inputPath = "ProgressReport.xlsx";
                const string outputPath = "ProgressReport_Gradient.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range of progress cells (e.g., A2:A20)
                Aspose.Cells.Range progressRange = worksheet.Cells.CreateRange("A2:A20");

                // Create a new style for the fill
                Style fillStyle = workbook.CreateStyle();

                // Apply a solid fill (green) as a fallback when gradient APIs are unavailable
                fillStyle.Pattern = BackgroundType.Solid;
                fillStyle.ForegroundColor = Color.Green;

                // Apply only the cell shading (no font or border changes)
                StyleFlag styleFlag = new StyleFlag
                {
                    CellShading = true
                };

                // Apply the style to the defined range
                progressRange.ApplyStyle(fillStyle, styleFlag);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
