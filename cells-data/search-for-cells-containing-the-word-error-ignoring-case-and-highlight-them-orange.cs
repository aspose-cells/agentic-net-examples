using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightError
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample data containing the word "error" in different cases
            worksheet.Cells["A1"].PutValue("Error occurred");
            worksheet.Cells["B2"].PutValue("no issues");
            worksheet.Cells["C3"].PutValue("critical ERROR detected");
            worksheet.Cells["D4"].PutValue("All good");

            // Create a style with orange background to highlight cells
            Style orangeStyle = workbook.CreateStyle();
            orangeStyle.ForegroundColor = Color.Orange;
            orangeStyle.Pattern = BackgroundType.Solid;

            // Configure find options: case‑insensitive, search within values, contains the text
            FindOptions findOptions = new FindOptions
            {
                CaseSensitive = false,
                LookInType = LookInType.Values,
                LookAtType = LookAtType.Contains
            };

            // Search and highlight all cells that contain the word "error"
            Cell previousCell = null;
            Cell foundCell;
            while ((foundCell = worksheet.Cells.Find("error", previousCell, findOptions)) != null)
            {
                foundCell.SetStyle(orangeStyle);
                // Continue searching from the cell after the current one
                previousCell = foundCell;
            }

            // Save the workbook
            workbook.Save("HighlightedErrors.xlsx");
        }
    }
}