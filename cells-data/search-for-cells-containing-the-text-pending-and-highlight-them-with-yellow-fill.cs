// Title: Highlight Excel cells that contain the word "Pending" with a yellow fill using Aspose.Cells for .NET
// AI Prompts: Search a worksheet for cells whose value includes "Pending" and apply a solid yellow background style with Aspose.Cells in C#. | Implement a case‑insensitive find operation for the text "Pending" and style all matching cells with a custom yellow fill using Aspose.Cells .NET API.
// Common Searches: How to locate cells that contain the word "Pending" and change their background color with Aspose.Cells in C# | C# Aspose.Cells example for searching text within a worksheet and applying a yellow fill | Using FindOptions to highlight cells with specific text in an Excel workbook via .NET | Apply conditional formatting programmatically for "Pending" status rows using Aspose.Cells | Search and style Excel cells based on value with Aspose.Cells for .NET
// Tags: find cells by substring Aspose.Cells | set cell fill color Aspose.Cells | highlight pending text Excel .NET | Aspose.Cells FindOptions usage | value‑based cell styling Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightPending
{
    // The code loads an existing workbook, creates a solid yellow style, configures FindOptions to locate cells whose values contain the string "Pending", iterates through each match applying the style, and saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path) or create a new one
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with yellow fill
            Style yellowStyle = workbook.CreateStyle();
            yellowStyle.ForegroundColor = Color.Yellow;
            yellowStyle.Pattern = BackgroundType.Solid;

            // Configure find options to search for cells that contain the text "Pending"
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                LookAtType = LookAtType.Contains          // Match if the cell value contains the text
            };

            // Iterate through all matching cells and apply the yellow style
            Cell previousCell = null;
            while (true)
            {
                Cell foundCell = worksheet.Cells.Find("Pending", previousCell, findOptions);
                if (foundCell == null)
                    break; // No more matches

                foundCell.SetStyle(yellowStyle);
                previousCell = foundCell; // Continue searching from the last found cell
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
