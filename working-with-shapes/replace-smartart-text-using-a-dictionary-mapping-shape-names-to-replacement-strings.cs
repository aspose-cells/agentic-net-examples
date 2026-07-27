// Title: Replace SmartArt Text in Excel with Aspose.Cells (C#) Using a Name‑to‑Text Dictionary
// Description: Loads an Excel workbook, creates a case‑insensitive dictionary that maps SmartArt shape names to new strings, iterates through all worksheets and shapes, detects SmartArt objects, accesses their grouped sub‑shapes via GetResultOfSmartArt().GetGroupedShapes(), updates the Text property of matching shapes, and saves the file with UpdateSmartArt enabled.
// Keywords: Aspose.Cells | C# | SmartArt | replace text | dictionary mapping | GetResultOfSmartArt | GetGroupedShapes | UpdateSmartArt | Excel automation | shape name lookup
// Common Searches: How to change SmartArt shape text in Excel using Aspose.Cells C# | C# replace specific SmartArt items by name | Aspose.Cells update SmartArt after editing text | Dictionary based text replacement for SmartArt in .xlsx
// Developer Intent: Update the text of selected SmartArt shapes in an Excel workbook based on a name‑to‑text dictionary.
// Use Cases: Generate quarterly reports by programmatically inserting titles and bullet points into SmartArt placeholders. | Localize Excel templates by swapping placeholder shape names with translated strings. | Populate presentation‑style spreadsheets where SmartArt elements are filled from a database or API at runtime.
// AI Prompts: Write C# code with Aspose.Cells that replaces SmartArt shape text using a dictionary and saves the workbook with UpdateSmartArt true. | Explain the purpose of GetResultOfSmartArt().GetGroupedShapes() and how to iterate over the returned shapes to modify their Text property. | Provide error‑handling strategies for missing shape names, unsupported SmartArt types, and save‑time failures when updating SmartArt text.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Saving;

// Loads an Excel workbook, creates a case‑insensitive dictionary that maps SmartArt shape names to new strings, iterates through all worksheets and shapes, detects SmartArt objects, accesses their grouped sub‑shapes via GetResultOfSmartArt().GetGroupedShapes(), updates the Text property of matching shapes, and saves the file with UpdateSmartArt enabled.
class SmartArtTextReplacer
{
    static void Main()
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the mapping from shape names to replacement texts
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "TitleShape", "Quarterly Report" },
            { "SubtitleShape", "Fiscal Year 2025" },
            { "Bullet1", "Increase sales by 15%" },
            { "Bullet2", "Launch new product line" }
        };

        // Iterate through all worksheets and their shapes
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Get the grouped shapes that represent the SmartArt layout
                    var grouped = shape.GetResultOfSmartArt().GetGroupedShapes();

                    foreach (Shape smartArtShape in grouped)
                    {
                        // If the shape's name exists in the dictionary, replace its text
                        if (replacements.TryGetValue(smartArtShape.Name, out string newText))
                        {
                            smartArtShape.Text = newText;
                        }
                    }
                }
            }
        }

        // Save the workbook with SmartArt update enabled (lifecycle rule: save)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.UpdateSmartArt = true;
        workbook.Save("output.xlsx", saveOptions);
    }
}
