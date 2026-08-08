// Title: C# AspNet Cells – Search for “error” (case‑insensitive) and apply orange highlight
// Description: Opens or creates an Excel workbook, sets up FindOptions to locate cells whose text contains the word “error” irrespective of case, assigns a solid orange fill to each match, and writes the updated file to HighlightedErrorCells.xlsx.
// Keywords: Aspose.Cells | C# | .NET | FindOptions | text search Excel | highlight cells orange | cell formatting | workbook automation | Excel styling | case insensitive search
// Common Searches: Aspose.Cells find text ignoring case | apply orange background to cells containing a keyword | C# loop through all matches in a worksheet | save workbook after conditional formatting with Aspose
// Developer Intent: Identify every cell that includes the word “error” regardless of letter case and color it orange.
// Use Cases: Automatically flag error messages in generated reports for quick visual review. | Run a validation routine on imported spreadsheets that highlights any occurrence of the word “error”. | Process a batch of workbooks, marking all error cells before archiving or sending to stakeholders.
// AI Prompts: Generate C# code using Aspose.Cells to locate all cells containing the substring “error” without case sensitivity and set their background to orange. | Show how to configure FindOptions with LookInType.Values and LookAtType.Contains for a case‑agnostic text search in Aspose.Cells. | Explain an efficient way to style each found cell only once, avoiding repeated style creation, in a .NET Excel automation script.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightError
{
    // Opens or creates an Excel workbook, sets up FindOptions to locate cells whose text contains the word “error” irrespective of case, assigns a solid orange fill to each match, and writes the updated file to HighlightedErrorCells.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Set up find options: case‑insensitive, search in cell values,
            // and look for cells that contain the word "error".
            // ------------------------------------------------------------
            FindOptions findOptions = new FindOptions
            {
                CaseSensitive = false,                 // ignore case
                LookInType = LookInType.Values,        // search cell values
                LookAtType = LookAtType.Contains        // match if the value contains the key
            };

            // ------------------------------------------------------------
            // Loop through all cells that match the criteria.
            // The previousCell parameter is used to continue the search.
            // ------------------------------------------------------------
            Cell previousCell = null;
            Cell foundCell;

            while ((foundCell = worksheet.Cells.Find("error", previousCell, findOptions)) != null)
            {
                // Create a style with orange background
                Style orangeStyle = workbook.CreateStyle();
                orangeStyle.ForegroundColor = Color.Orange;
                orangeStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the found cell
                foundCell.SetStyle(orangeStyle);

                // Update previousCell to continue searching from the next cell
                previousCell = foundCell;
            }

            // ------------------------------------------------------------
            // Save the workbook with the highlighted cells
            // ------------------------------------------------------------
            workbook.Save("HighlightedErrorCells.xlsx");
        }
    }
}
