// Title: Highlight 'Pending' cells with yellow fill using Aspose.Cells for .NET (C#)
// Description: Loads a workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate every cell whose value contains the word "Pending", creates a solid yellow style, applies it to each matching cell, and saves the updated file.
// Keywords: Aspose.Cells | C# | .NET | Excel | find cells containing text | highlight cells | yellow fill | FindOptions | LookInType.Values | LookAtType.Contains | search pending cells | apply cell style | worksheet
// Common Searches: Aspose.Cells find text and highlight cells | C# highlight cells containing Pending | How to search for a string in Excel with Aspose.Cells | Apply background color to cells using Aspose.Cells .NET | Find and style cells by keyword in C#
// Developer Intent: Search the worksheet for every cell that includes the word "Pending" and apply a yellow background style to it.
// Use Cases: Flag pending tasks in a project‑tracking spreadsheet. | Highlight rows awaiting approval in status reports. | Mark pending invoices or payments in accounting sheets. | Identify incomplete entries during data validation.
// AI Prompts: Generate C# code with Aspose.Cells that finds all cells containing "Pending" and sets their fill color to yellow. | Explain how to make the search case‑insensitive and reuse a single Style object for better performance. | Show how to also change the font color and add a comment to the highlighted cells. | Provide an example that processes every worksheet in a workbook to highlight "Pending" cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace HighlightPendingCells
{
    // Loads a workbook, uses FindOptions (LookInType.Values, LookAtType.Contains) to locate every cell whose value contains the word "Pending", creates a solid yellow style, applies it to each matching cell, and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path) or create a new one
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure find options to search for cells whose value contains "Pending"
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,      // Search in cell values
                LookAtType = LookAtType.Contains      // Match if the value contains the text
            };

            // Find the first occurrence
            Cell foundCell = worksheet.Cells.Find("Pending", null, findOptions);

            // Loop through all matching cells
            while (foundCell != null)
            {
                // Create a style with yellow fill
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the found cell
                foundCell.SetStyle(highlightStyle);

                // Find the next occurrence, starting after the current cell
                foundCell = worksheet.Cells.Find("Pending", foundCell, findOptions);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
