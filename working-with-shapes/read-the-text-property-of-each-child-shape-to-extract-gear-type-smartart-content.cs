// Title: Extract Gear SmartArt Child Shape Text in Excel using Aspose.Cells for .NET
// Description: Loads an Excel workbook, scans every worksheet for SmartArt objects, retrieves the grouped result with GetResultOfSmartArt, reads the Text property of each child shape, and prints the worksheet name, SmartArt ID, child shape ID, and its text. The workbook can then be saved optionally.
// Keywords: Aspose.Cells SmartArt extraction | Gear SmartArt child text C# | GetResultOfSmartArt example | read shape text Excel .NET | iterate Excel shapes Aspose
// Common Searches: how to read child shape text from SmartArt using Aspose.Cells | extract Gear SmartArt labels in C# | Aspose.Cells GetResultOfSmartArt tutorial | list all SmartArt shapes and their texts in an Excel file | C# code to iterate Excel shapes and get text
// Developer Intent: Retrieve the Text property of each child shape inside a Gear‑type SmartArt object in an Excel workbook.
// Use Cases: Generate a report of all SmartArt labels with their worksheet and shape identifiers. | Collect SmartArt captions for import into a database or documentation system. | Validate that required SmartArt elements contain expected text before further processing.
// AI Prompts: Create a method that returns a dictionary mapping each worksheet to a list of child shape texts for every SmartArt object. | Show how to filter SmartArt shapes by name or ID and extract child texts that match a specific regular expression. | Write code to export SmartArt child shape texts, along with worksheet and shape IDs, to a CSV file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtReader
{
    // Loads an Excel workbook, scans every worksheet for SmartArt objects, retrieves the grouped result with GetResultOfSmartArt, reads the Text property of each child shape, and prints the worksheet name, SmartArt ID, child shape ID, and its text. The workbook can then be saved optionally.
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithSmartArt.xlsx";
            const string outputPath = "Output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through worksheets and shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Process only SmartArt shapes
                        if (shape.IsSmartArt)
                        {
                            GroupShape group = shape.GetResultOfSmartArt();
                            if (group != null)
                            {
                                foreach (Shape child in group.GetGroupedShapes())
                                {
                                    string childText = child.Text;
                                    Console.WriteLine($"Worksheet: {sheet.Name}, SmartArt Shape ID: {shape.Id}, Child Shape ID: {child.Id}, Text: {childText}");
                                }
                            }
                        }
                    }
                }

                // Save the workbook (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
