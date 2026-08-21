// Title: Validate non‑empty SmartArt node text after replacement using Aspose.Cells for .NET
// Description: C# sample that opens an Excel workbook, walks through each worksheet and shape, converts SmartArt shapes to GroupShape objects, overwrites the text of every SmartArt node, verifies the new value is not blank, logs the outcome, and saves the file with UpdateSmartArt enabled so the changes are persisted.
// Keywords: Aspose.Cells | SmartArt validation | C# Excel | GroupShape | UpdateSmartArt | replace SmartArt text | non‑empty text check | workbook processing .NET | Excel SmartArt node iteration | Aspose.Cells example
// Common Searches: how to check SmartArt node text after modification with Aspose.Cells | C# validate SmartArt shapes are not empty | Aspose.Cells replace SmartArt text and verify | save Excel workbook with updated SmartArt using Aspose | iterate SmartArt nodes in .NET
// Developer Intent: Ensure every SmartArt node in an Excel file contains text after a bulk replacement operation.
// Use Cases: Batch update placeholder text in SmartArt diagrams and confirm each node retains content. | Log identifiers of SmartArt nodes that become empty during automated processing. | Persist SmartArt edits by saving the workbook with the UpdateSmartArt option. | Integrate SmartArt validation into CI pipelines that generate Excel reports.
// AI Prompts: Write C# code that throws an exception if any SmartArt node text is empty after replacement using Aspose.Cells. | Show how to collect IDs of SmartArt nodes with blank text into a list for later reporting. | Create a method that conditionally replaces SmartArt node text and returns a boolean indicating overall validation success.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartArtValidation
{
    // C# sample that opens an Excel workbook, walks through each worksheet and shape, converts SmartArt shapes to GroupShape objects, overwrites the text of every SmartArt node, verifies the new value is not blank, logs the outcome, and saves the file with UpdateSmartArt enabled so the changes are persisted.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all shapes in the worksheet
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert SmartArt to a group of shapes
                        GroupShape groupShape = shape.GetResultOfSmartArt();

                        // Iterate through each grouped shape (individual SmartArt nodes)
                        foreach (Shape smartArtNode in groupShape.GetGroupedShapes())
                        {
                            // Example replacement: set some text (you can customize this)
                            smartArtNode.Text = "ReplacedText";

                            // Validate that the text is not empty after replacement
                            if (string.IsNullOrWhiteSpace(smartArtNode.Text))
                            {
                                Console.WriteLine($"Validation failed: SmartArt node (Id={smartArtNode.Id}) has empty text.");
                            }
                            else
                            {
                                Console.WriteLine($"SmartArt node (Id={smartArtNode.Id}) text validated: \"{smartArtNode.Text}\"");
                            }
                        }
                    }
                }
            }

            // Save the workbook with SmartArt update enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // Ensure SmartArt changes are persisted
            workbook.Save("output.xlsx", saveOptions);
        }
    }
}
