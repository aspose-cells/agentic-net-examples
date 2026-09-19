// Title: Load an XLSX workbook with Aspose.Cells for .NET and enumerate shapes to detect and (where possible) replace SmartArt text
// AI Prompts: Generate C# code that opens a .xlsx file using Aspose.Cells, iterates through all worksheets, finds shapes where IsSmartArt is true, and replaces each SmartArt node's text with a supplied string, providing a fallback notice when the API version lacks SmartArt support. | Show how to list every SmartArt shape in an Excel workbook with Aspose.Cells for .NET, outputting the worksheet name and shape identifier for each detected SmartArt object. | Create a method that saves the workbook after attempting SmartArt text replacement, ensures the output directory exists, and includes robust exception handling for file I/O and API limitations.
// Common Searches: C# Aspose.Cells replace text in SmartArt shapes of an Excel file | How to enumerate worksheet shapes and identify SmartArt using Aspose.Cells .NET | Aspose.Cells IsSmartArt property example code | SmartArt text modification in XLSX with Aspose.Cells version limitations | Load Excel workbook and update SmartArt nodes using Aspose.Cells for .NET
// Tags: Aspose.Cells replace SmartArt text C# | enumerate worksheet shapes Aspose.Cells | detect SmartArt shapes Excel Aspose.Cells | load and save XLSX workbook Aspose.Cells .NET | SmartArt manipulation limitation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an input.xlsx workbook with Aspose.Cells, checks that the file exists, iterates over each worksheet and its ShapeCollection, identifies shapes flagged as SmartArt via the IsSmartArt property, logs a message for each SmartArt shape noting that current API versions do not support direct manipulation, ensures the output directory is present, saves the workbook as output.xlsx, and includes a placeholder recursive method for SmartArt node text replacement that can be activated when a supporting API version becomes available.
class ReplaceSmartArtText
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get all shapes on the sheet
                ShapeCollection shapes = sheet.Shapes;

                // Process each shape
                foreach (Shape shape in shapes)
                {
                    // Identify SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Aspose.Cells version used does not expose a SmartArt API.
                        // If a newer version is available, replace this block with proper SmartArt handling.
                        Console.WriteLine($"SmartArt shape found on sheet \"{sheet.Name}\" but SmartArt manipulation is not supported in this version.");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder for SmartArt node text replacement (not used with current API version)
    private static void ReplaceTextInSmartArtNode(dynamic node)
    {
        if (node == null) return;

        try
        {
            // Update the text of the current node if a TextFrame exists
            if (node.TextFrame != null)
            {
                node.TextFrame.Text = "New Text";
            }

            // Process child nodes
            foreach (var child in node.ChildNodes)
            {
                ReplaceTextInSmartArtNode(child);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while replacing text in SmartArt node: {ex.Message}");
        }
    }
}
