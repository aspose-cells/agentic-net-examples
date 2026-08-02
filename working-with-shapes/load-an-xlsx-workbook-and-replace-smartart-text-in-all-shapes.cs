// Title: Replace SmartArt Text in All Shapes of an XLSX Workbook with Aspose.Cells for .NET
// Description: This C# example loads an XLSX file using Aspose.Cells, scans every worksheet for SmartArt shapes, converts each SmartArt to a GroupShape, overwrites the text of every grouped element with a new value, and saves the workbook with UpdateSmartArt enabled so the changes are persisted.
// Keywords: Aspose.Cells SmartArt text replacement | C# update Excel SmartArt | Iterate workbook shapes Aspose | OoxmlSaveOptions UpdateSmartArt | Bulk edit SmartArt captions | .NET Excel shape manipulation
// Common Searches: how to change SmartArt text in Excel using Aspose.Cells | Aspose.Cells replace all SmartArt labels C# | programmatically edit SmartArt shapes in .xlsx | save Excel file after modifying SmartArt Aspose | bulk update SmartArt captions across worksheets
// Developer Intent: Programmatically replace the text of every SmartArt shape in an Excel workbook and write the changes back to the file.
// Use Cases: Localize SmartArt labels in a template before distribution. | Populate SmartArt descriptions with dynamic report data. | Enforce brand‑consistent wording across multiple worksheets.
// AI Prompts: Generate C# code that iterates through a workbook, finds SmartArt shapes, and sets their text from a dictionary using Aspose.Cells. | Explain what happens if OoxmlSaveOptions.UpdateSmartArt is set to false when saving after SmartArt modifications. | Show how to log original SmartArt text before replacement while processing with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This C# example loads an XLSX file using Aspose.Cells, scans every worksheet for SmartArt shapes, converts each SmartArt to a GroupShape, overwrites the text of every grouped element with a new value, and saves the workbook with UpdateSmartArt enabled so the changes are persisted.
class ReplaceSmartArtText
{
    static void Main()
    {
        // Paths to the source and destination files
        string inputFile = "input.xlsx";
        string outputFile = "output.xlsx";

        // Load the workbook (lifecycle: load)
        Workbook workbook = new Workbook(inputFile);

        // Iterate through all worksheets and their shapes
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Convert SmartArt to a grouped shape collection
                    GroupShape group = shape.GetResultOfSmartArt();

                    // Replace the text of each grouped shape
                    foreach (Shape smartShape in group.GetGroupedShapes())
                    {
                        smartShape.Text = "ReplacedText";
                    }
                }
            }
        }

        // Save the workbook with SmartArt updating enabled (lifecycle: save)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save(outputFile, saveOptions);
    }
}
