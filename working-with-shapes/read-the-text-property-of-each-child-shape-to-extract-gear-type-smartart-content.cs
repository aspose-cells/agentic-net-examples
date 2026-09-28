// Title: Extract text from SmartArt shapes in every worksheet of an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a given .xlsx file with Aspose.Cells, verifies the file exists, loops through all worksheets and their shapes, and returns a list of the Text property for each SmartArt shape while handling shape‑level exceptions. | Create a reusable C# method that accepts a workbook path, uses Aspose.Cells to collect non‑empty SmartArt text from all worksheets, and prints the extracted strings to the console with proper error handling.
// Common Searches: how to read smartart text from an Excel workbook using Aspose.Cells C# | c# Aspose.Cells iterate shapes and get smartart content | extract all smartart shape texts from multiple worksheets in .xlsx with Aspose.Cells | Aspose.Cells get Text property of SmartArt shapes in .NET | C# sample code to list smartart texts in Excel file using Aspose.Cells
// Tags: extract smartart text Aspose.Cells | iterate worksheet shapes Aspose.Cells | smartart shape text .NET | read smartart content from xlsx | file existence check Aspose.Cells | shape level error handling Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Collections.Generic;
using System.IO;

// The example verifies that an input.xlsx file exists, loads it with Aspose.Cells, iterates each worksheet and every shape, extracts the Text property of shapes identified as SmartArt, collects non‑empty texts into a list, and writes each extracted string to the console, with shape‑level exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold extracted SmartArt texts
            List<string> smartArtTexts = new List<string>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Process only SmartArt shapes
                        if (shape.IsSmartArt)
                        {
                            // Retrieve the text associated with the SmartArt shape
                            string text = shape.Text;
                            if (!string.IsNullOrEmpty(text))
                            {
                                smartArtTexts.Add(text);
                            }
                        }
                    }
                    catch (Exception exShape)
                    {
                        Console.WriteLine($"Error processing shape: {exShape.Message}");
                    }
                }
            }

            // Output the extracted texts
            foreach (string txt in smartArtTexts)
            {
                Console.WriteLine(txt);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
