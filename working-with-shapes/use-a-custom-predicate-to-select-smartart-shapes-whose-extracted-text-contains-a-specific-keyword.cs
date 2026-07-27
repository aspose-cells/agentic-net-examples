// Title: Select SmartArt Shapes by Keyword Using a Custom Predicate in Aspose.Cells for .NET
// Description: Shows how to load an Excel workbook with Aspose.Cells, create a case‑insensitive predicate that flags SmartArt shapes whose Text contains a target keyword, walk through all worksheets, gather matching SmartArtShape objects, print their Name and Text, and optionally save the file.
// Keywords: Aspose.Cells | C# SmartArt filter | SmartArt shape predicate | search SmartArt text | filter shapes by keyword | Excel SmartArt extraction | case‑insensitive text search | shape.IsSmartArt | SmartArtShape collection | Aspose.Cells API
// Common Searches: How to filter SmartArt shapes by text in Aspose.Cells | C# find SmartArt containing a word in an Excel workbook | Aspose.Cells select shapes with specific keyword | Retrieve SmartArt objects based on text content | Use predicate to locate SmartArt in .NET
// Developer Intent: Locate and return every SmartArt shape whose displayed text includes a specified keyword.
// Use Cases: Generate a list of SmartArt items that mention a product name for reporting. | Remove or replace SmartArt graphics that contain prohibited terminology. | Create a summary sheet that aggregates SmartArt names and their texts across multiple worksheets. | Automate quality checks by flagging SmartArt objects with missing or incorrect keywords.
// AI Prompts: Write C# code with Aspose.Cells that replaces a given keyword in the Text property of all matching SmartArt shapes. | Show how to group the found SmartArt shapes by worksheet and export each group to a separate Excel file. | Provide an example that makes the predicate case‑sensitive and logs the names of shapes that do not contain the keyword.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to load an Excel workbook with Aspose.Cells, create a case‑insensitive predicate that flags SmartArt shapes whose Text contains a target keyword, walk through all worksheets, gather matching SmartArtShape objects, print their Name and Text, and optionally save the file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Keyword to search within SmartArt text
        string keyword = "TargetKeyword";

        // Predicate that returns true for SmartArt shapes whose text contains the keyword (case‑insensitive)
        Func<Shape, bool> smartArtWithKeyword = shape =>
            shape.IsSmartArt &&
            !string.IsNullOrEmpty(shape.Text) &&
            shape.Text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

        // List to hold the matching SmartArt shapes
        List<SmartArtShape> matchingSmartArts = new List<SmartArtShape>();

        // Iterate through all worksheets and their shapes
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Shape shape in sheet.Shapes)
            {
                if (smartArtWithKeyword(shape))
                {
                    // Cast to SmartArtShape (safe because IsSmartArt is true)
                    matchingSmartArts.Add((SmartArtShape)shape);
                }
            }
        }

        // Example action: print the names and texts of the found SmartArt shapes
        foreach (SmartArtShape smartArt in matchingSmartArts)
        {
            Console.WriteLine($"Found SmartArt - Name: {smartArt.Name}, Text: {smartArt.Text}");
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}
