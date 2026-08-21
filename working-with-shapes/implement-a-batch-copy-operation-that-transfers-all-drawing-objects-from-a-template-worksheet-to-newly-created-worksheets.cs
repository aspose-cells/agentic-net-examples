// Title: Batch copy drawing objects from a template worksheet to multiple new sheets with Aspose.Cells for .NET
// Description: Loads a template workbook, treats its first worksheet as the source of drawing objects, creates several new worksheets, and copies each shape (pictures, charts, text boxes, etc.) to the new sheets while preserving position and size, then saves the result.
// Keywords: Aspose.Cells copy shapes | C# batch copy drawing objects | AddCopy shape Aspose.Cells | duplicate worksheet graphics .NET | copy charts and pictures between sheets | automate workbook layout Aspose | template worksheet shapes | copy all shapes Aspose.Cells
// Common Searches: copy all shapes from one worksheet to many sheets Aspose.Cells | batch copy drawing objects to new worksheets .NET | Aspose.Cells duplicate charts and images across sheets | how to use AddCopy for shapes in Aspose.Cells | C# copy template worksheet graphics to multiple sheets
// Developer Intent: Copy every shape from a template worksheet to newly created worksheets within the same workbook.
// Use Cases: Apply a common logo, header, and chart template to each section of a multi‑sheet report. | Generate client‑specific worksheets that share identical visual elements without manual recreation. | Automate the cloning of a design‑rich template sheet when building data‑driven workbooks with repeated layouts.
// AI Prompts: Show C# code that copies all shapes from a source worksheet to all existing worksheets using Aspose.Cells. | Add comprehensive error handling for missing template files and unsupported shape types while copying. | Explain how to offset the position of copied shapes on the target worksheets after using AddCopy.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchCopyDrawingObjects
{
    // Loads a template workbook, treats its first worksheet as the source of drawing objects, creates several new worksheets, and copies each shape (pictures, charts, text boxes, etc.) to the new sheets while preserving position and size, then saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the template workbook and the result workbook
                string templatePath = "Template.xlsx";
                string outputPath = "Result.xlsx";

                // Ensure the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                }

                // Load the template workbook (contains the source worksheet with drawing objects)
                Workbook workbook = new Workbook(templatePath);

                // Assume the first worksheet is the template sheet
                Worksheet templateSheet = workbook.Worksheets[0];

                // Names of the new worksheets that will receive the copied drawing objects
                List<string> newSheetNames = new List<string> { "Copy1", "Copy2", "Copy3" };

                foreach (string sheetName in newSheetNames)
                {
                    // Add a new worksheet to the workbook and obtain the worksheet reference
                    Worksheet newSheet = workbook.Worksheets.Add(sheetName);

                    // Copy each shape (including pictures, charts, etc.) from the template sheet
                    foreach (Shape sourceShape in templateSheet.Shapes)
                    {
                        // Copy the shape preserving its original position and size
                        newSheet.Shapes.AddCopy(
                            sourceShape,
                            sourceShape.UpperLeftRow,
                            sourceShape.UpperLeftColumn,
                            sourceShape.LowerRightRow,
                            sourceShape.LowerRightColumn);
                    }
                }

                // Save the workbook with the newly created sheets and copied drawing objects
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
