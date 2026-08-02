// Title: C# – Batch replace SmartArt text across all worksheets using Aspose.Cells
// Description: Loads an Excel file, walks through every worksheet and shape, extracts SmartArt via GetResultOfSmartArt, swaps a target placeholder with new text in each inner shape, and saves the workbook with OoxmlSaveOptions.UpdateSmartArt to apply the changes.
// Keywords: Aspose.Cells SmartArt text replace | C# batch update SmartArt | GetResultOfSmartArt | UpdateSmartArt flag | iterate worksheets Aspose.Cells | GroupShape inner text | Excel SmartArt automation | Aspose.Cells .NET example
// Common Searches: replace text in all SmartArt shapes Aspose.Cells C# | how to update SmartArt across multiple sheets | Aspose.Cells batch modify SmartArt diagram | save workbook with updated SmartArt
// Developer Intent: Automatically substitute a specific placeholder string with new content in every SmartArt diagram present in all worksheets of an Excel workbook.
// Use Cases: Refresh company branding in SmartArt diagrams of a multi‑sheet financial report. | Swap placeholder labels in training templates that rely on SmartArt flowcharts. | Update product names in regional sales decks that contain SmartArt visuals.
// AI Prompts: Write C# code that finds and replaces a given phrase in all SmartArt objects of an Excel workbook using Aspose.Cells. | Explain the purpose of OoxmlSaveOptions.UpdateSmartArt and show how to enable it when saving a workbook. | Describe step‑by‑step how GetResultOfSmartArt converts a SmartArt shape to a GroupShape for text manipulation.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartArtBatchReplace
{
    // Loads an Excel file, walks through every worksheet and shape, extracts SmartArt via GetResultOfSmartArt, swaps a target placeholder with new text in each inner shape, and saves the workbook with OoxmlSaveOptions.UpdateSmartArt to apply the changes.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains SmartArt objects
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Text to find and its replacement
            string oldText = "Placeholder";
            string newText = "Replaced Text";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all shapes on the worksheet
                foreach (Shape shape in sheet.Shapes)
                {
                    // Process only SmartArt shapes
                    if (shape.IsSmartArt)
                    {
                        // Convert the SmartArt to its grouped shape representation
                        GroupShape smartArtGroup = shape.GetResultOfSmartArt();

                        // Iterate through each individual shape that makes up the SmartArt
                        foreach (Shape innerShape in smartArtGroup.GetGroupedShapes())
                        {
                            // Replace the text inside the shape if it contains the old text
                            if (!string.IsNullOrEmpty(innerShape.Text) && innerShape.Text.Contains(oldText))
                            {
                                innerShape.Text = innerShape.Text.Replace(oldText, newText);
                            }
                        }
                    }
                }
            }

            // Save the workbook with SmartArt update enabled
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateSmartArt = true; // Apply the text changes to SmartArt
            workbook.Save("OutputWorkbook.xlsx", saveOptions);
        }
    }
}
