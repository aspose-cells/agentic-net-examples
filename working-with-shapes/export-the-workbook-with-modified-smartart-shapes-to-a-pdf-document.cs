// Title: Export an Excel workbook with SmartArt shapes to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel file that contains SmartArt to PDF | How to save a workbook with SmartArt diagrams as PDF in .NET | Iterate through worksheet shapes before converting to PDF with Aspose.Cells | Convert Excel workbook with graphics and SmartArt to PDF using Aspose.Cells | Save Excel workbook as PDF while preserving SmartArt objects in C#
// Tags: aspocells saveformat.pdf conversion | iterate worksheet shapes c# | preserve smartart during pdf export | excel to pdf with graphics c# | export workbook to pdf aspocells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, walks through each worksheet and its shapes (including SmartArt), ensures the output folder exists, and then saves the workbook as a PDF using SaveFormat.Pdf, with comprehensive exception handling.
    class Program
    {
        static void Main()
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all shapes on the worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            // Example: handle picture shapes
                            // Note: The Shape.Type enum may vary between Aspose.Cells versions.
                            // If needed, add version‑specific handling here.
                            // For now we simply acknowledge the shape without type checking.
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"Error processing shape on sheet '{sheet.Name}': {shapeEx.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the workbook to PDF
                workbook.Save(outputFile, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
