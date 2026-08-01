// Title: C# – Batch copy all drawing objects from a template worksheet to multiple new sheets with Aspose.Cells
// Description: Loads a template workbook, extracts every Shape from the "Template" worksheet, creates a destination workbook, adds several blank worksheets, and uses Shapes.AddCopy to duplicate each drawing object on every new sheet while preserving its original row and column placement. The final workbook is saved as XLSX.
// Keywords: Aspose.Cells | C# copy shapes | batch copy drawing objects | worksheet shapes | AddCopy method | Aspose.Cells .NET example | duplicate drawing objects | copy shapes multiple sheets | Excel shape cloning | Aspose.Cells GitHub sample | Aspose.Cells API | copy charts programmatically
// Common Searches: Aspose.Cells copy shapes to new worksheet C# | batch copy drawing objects Aspose.Cells .NET | duplicate all shapes from a template sheet using Aspose | how to copy worksheet shapes to multiple sheets in C# | AddCopy shape example Aspose.Cells | copy charts and images across worksheets C# Aspose | Aspose.Cells shape cloning code sample
// Developer Intent: Programmatically duplicate every shape from a template worksheet onto each newly created worksheet in a separate workbook.
// Use Cases: Generate a report workbook where a pre‑designed header graphic is automatically placed on every report sheet. | Create a batch of invoice worksheets that share the same company logo, watermark, and decorative shapes. | Clone a dashboard layout—including charts, images, and connectors—across several analysis sheets while keeping exact positions.
// AI Prompts: Write C# code using Aspose.Cells to copy all shapes from a source worksheet to a list of target worksheets, preserving row/column positions and handling errors. | Explain how to modify the batch copy loop to apply different row or column offsets based on the target worksheet index. | Show how to extend the example to also copy shape formatting such as line style, fill color, and text formatting.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsBatchCopyDrawingObjects
{
    // Loads a template workbook, extracts every Shape from the "Template" worksheet, creates a destination workbook, adds several blank worksheets, and uses Shapes.AddCopy to duplicate each drawing object on every new sheet while preserving its original row and column placement. The final workbook is saved as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "TemplateWorkbook.xlsx";
                const string destinationPath = "DestinationWorkbook.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook that contains the source worksheet with drawing objects
                Workbook templateWorkbook = new Workbook(templatePath);
                Worksheet templateSheet = templateWorkbook.Worksheets["Template"]; // source sheet name

                if (templateSheet == null)
                {
                    Console.WriteLine("Source worksheet 'Template' not found in the template workbook.");
                    return;
                }

                // Create a new workbook that will receive the copied drawing objects
                Workbook destinationWorkbook = new Workbook();
                // Remove the default sheet created by the constructor
                destinationWorkbook.Worksheets.Clear();

                // Names of the new worksheets to be created
                string[] newSheetNames = { "CopySheet1", "CopySheet2", "CopySheet3" };

                foreach (string sheetName in newSheetNames)
                {
                    // Add a new blank worksheet
                    Worksheet newSheet = destinationWorkbook.Worksheets.Add(sheetName);

                    // Iterate through all shapes (drawing objects) in the template sheet
                    foreach (Shape sourceShape in templateSheet.Shapes)
                    {
                        try
                        {
                            // Copy each shape to the new worksheet preserving its original cell position.
                            // Offsets are set to 0 because the Shape class in this version does not expose offset properties.
                            newSheet.Shapes.AddCopy(
                                sourceShape,
                                sourceShape.UpperLeftRow,
                                0, // row offset (pixels)
                                sourceShape.UpperLeftColumn,
                                0  // column offset (pixels)
                            );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to copy shape '{sourceShape.Name}': {ex.Message}");
                        }
                    }
                }

                // Save the result workbook
                destinationWorkbook.Save(destinationPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {destinationPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
