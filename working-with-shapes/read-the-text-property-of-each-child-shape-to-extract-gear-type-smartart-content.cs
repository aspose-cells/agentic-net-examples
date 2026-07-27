// Title: C# – Extract Gear6 & Gear9 SmartArt Text with Aspose.Cells
// Description: Loads an Excel workbook, walks through each worksheet and shape, converts SmartArt objects to GroupShape, then reads the Text property of every child shape whose AutoShapeType is Gear6 or Gear9.
// Keywords: Aspose.Cells | C# | SmartArt extraction | Gear6 | Gear9 | AutoShapeType | grouped shapes | Excel text retrieval | shape text read
// Common Searches: how to read gear shape text in smartart using aspose.cells | c# extract smartart child shapes text | convert smartart to groupshape aspnet | retrieve gear6 gear9 labels from excel
// Developer Intent: Read the Text property of each Gear6 or Gear9 child shape inside SmartArt objects in an Excel file.
// Use Cases: Generate a list of gear labels from a process‑flow SmartArt for documentation. | Validate that every gear icon in a maintenance diagram contains required annotation before publishing. | Create a summary report of equipment identifiers by collecting gear shape text.
// AI Prompts: Write C# code with Aspose.Cells that returns a dictionary of Gear6 and Gear9 shape texts from all SmartArt in a workbook. | Add comprehensive error handling to the SmartArt extraction sample for missing files, empty groups, and null text values. | Show how to modify the extracted gear text and write the updated values back to the same shapes using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, walks through each worksheet and shape, converts SmartArt objects to GroupShape, then reads the Text property of every child shape whose AutoShapeType is Gear6 or Gear9.
class ExtractGearSmartArt
{
    static void Main()
    {
        string filePath = "SmartArt.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        try
        {
            // Load the workbook that contains the SmartArt
            Workbook workbook = new Workbook(filePath);

            // Loop through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Loop through each shape on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt to a GroupShape
                        GroupShape group = shape.GetResultOfSmartArt();
                        if (group != null)
                        {
                            // Iterate over the child shapes inside the group
                            foreach (Shape child in group.GetGroupedShapes())
                            {
                                // Identify gear shapes (Gear6 or Gear9)
                                if (child.AutoShapeType == AutoShapeType.Gear6 ||
                                    child.AutoShapeType == AutoShapeType.Gear9)
                                {
                                    // Output the text of the gear shape
                                    string gearText = child.Text;
                                    Console.WriteLine($"Gear ({child.AutoShapeType}) text: {gearText}");
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
