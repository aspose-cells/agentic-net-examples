// Title: Extract Gear SmartArt text from Excel worksheet shapes using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read text of SmartArt shapes in an Excel workbook | how to filter Excel shapes containing the word Gear with Aspose.Cells | iterate through worksheet ShapeCollection to get SmartArt content using Aspose.Cells .NET | extract specific SmartArt keyword from Excel file programmatically in C#
// Tags: Aspose.Cells read shape text .NET | SmartArt text extraction from Excel using C# | filter Gear SmartArt shapes Aspose.Cells | iterate worksheet shapes Aspose.Cells API | extract SmartArt keyword from workbook

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook with Aspose.Cells, iterates each worksheet's ShapeCollection, reads the Text property of every shape, collects texts that contain the word "Gear" (case‑insensitive), and outputs the extracted Gear SmartArt content to the console.
class GearSmartArtExtractor
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // List to hold extracted Gear SmartArt texts
        List<string> gearTexts = new List<string>();

        try
        {
            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Iterate through each shape
                foreach (Shape shape in shapes)
                {
                    try
                    {
                        // Attempt to read the Text property of the shape
                        string text = shape.Text;

                        // Simple heuristic: Gear SmartArt often contains the word "Gear"
                        if (!string.IsNullOrEmpty(text) &&
                            text.IndexOf("Gear", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            gearTexts.Add(text);
                        }
                    }
                    catch (Exception shapeEx)
                    {
                        // Log shape-specific errors but continue processing other shapes
                        Console.WriteLine($"Error processing shape: {shapeEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return;
        }

        // Output the extracted Gear SmartArt content to the console
        Console.WriteLine("Extracted Gear SmartArt Texts:");
        foreach (string txt in gearTexts)
        {
            Console.WriteLine(txt);
        }
    }
}
