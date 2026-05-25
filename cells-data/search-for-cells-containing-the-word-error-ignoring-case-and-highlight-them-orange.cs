using System;
using System.Drawing;
using Aspose.Cells;

class HighlightErrorCells
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // replace with new Workbook("input.xlsx") to load
        Worksheet worksheet = workbook.Worksheets[0];

        // Example data – in real scenario the workbook would already contain data
        worksheet.Cells["A1"].PutValue("No error here");
        worksheet.Cells["A2"].PutValue("Error found");
        worksheet.Cells["A3"].PutValue("Another ERROR case");
        worksheet.Cells["A4"].PutValue("All good");

        // Prepare find options: case‑insensitive, search in values, look for containing text
        FindOptions findOptions = new FindOptions
        {
            CaseSensitive = false,
            LookInType = LookInType.Values,
            LookAtType = LookAtType.Contains
        };

        // Style to apply (orange background)
        Style orangeStyle = workbook.CreateStyle();
        orangeStyle.ForegroundColor = Color.Orange;
        orangeStyle.Pattern = BackgroundType.Solid;

        // Search and highlight all matching cells
        Cell previousCell = null;
        Cell foundCell = worksheet.Cells.Find("error", previousCell, findOptions);
        while (foundCell != null)
        {
            foundCell.SetStyle(orangeStyle);
            // Continue searching from the cell just found
            previousCell = foundCell;
            foundCell = worksheet.Cells.Find("error", previousCell, findOptions);
        }

        // Save the workbook
        workbook.Save("HighlightedErrors.xlsx");
    }
}