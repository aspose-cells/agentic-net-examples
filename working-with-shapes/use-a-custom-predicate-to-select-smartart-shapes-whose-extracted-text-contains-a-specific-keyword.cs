// Title: Filter Excel SmartArt shapes by keyword using Aspose.Cells for .NET with a custom predicate
// Common Searches: Aspose.Cells C# filter SmartArt shapes that contain a specific word | how to extract SmartArt text and search for a keyword with Aspose.Cells | select SmartArt objects in an Excel workbook using a custom predicate in C# | retrieve names of SmartArt shapes containing "Important" using Aspose.Cells | dynamic GetSmartArt usage for keyword search in Excel file C#
// Tags: filter SmartArt shapes by text Aspose.Cells | dynamic GetSmartArt C# | extract all SmartArt text Aspose.Cells | keyword search in Excel SmartArt | iterate worksheet shapes Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Microsoft.CSharp.RuntimeBinder;

// The example loads an Excel workbook, iterates through each worksheet's shapes, dynamically accesses SmartArt objects, extracts their combined text, and adds shapes whose text contains a specified keyword (case‑insensitive) to a list. It then prints the matching shape names and saves the workbook.
class SmartArtSelector
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Keyword to search for inside SmartArt text
            string keyword = "Important";

            // List to hold shapes that match the predicate
            List<Shape> matchingShapes = new List<Shape>();

            // Iterate through all worksheets and their shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Shape shape in sheet.Shapes)
                {
                    try
                    {
                        // Use dynamic to call GetSmartArt at runtime (avoids compile‑time dependency)
                        dynamic dShape = shape;
                        dynamic smartArt = dShape.GetSmartArt();

                        if (smartArt != null)
                        {
                            // Retrieve all text from the SmartArt object
                            string allText = smartArt.GetAllText() as string;

                            // Check if the text contains the keyword (case‑insensitive)
                            if (!string.IsNullOrEmpty(allText) &&
                                allText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                matchingShapes.Add(shape);
                            }
                        }
                    }
                    catch (RuntimeBinderException)
                    {
                        // Shape does not support SmartArt – ignore and continue
                    }
                    catch (Exception ex)
                    {
                        // Log unexpected errors for the current shape
                        Console.WriteLine($"Error processing shape '{shape.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Output the names of matching SmartArt shapes
            Console.WriteLine($"SmartArt shapes containing the keyword \"{keyword}\":");
            foreach (Shape matchedShape in matchingShapes)
            {
                Console.WriteLine($"- Shape Name: {matchedShape.Name}, Sheet: {matchedShape.Worksheet.Name}");
            }

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors in the overall process
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
