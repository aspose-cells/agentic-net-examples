using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightPending
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data (optional – in real scenario the workbook would already contain data)
            cells["A1"].PutValue("Pending");
            cells["B2"].PutValue("Completed");
            cells["C3"].PutValue("Pending Review");
            cells["D4"].PutValue("Pending");

            // Configure find options to search for cells whose value contains "Pending"
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                LookAtType = LookAtType.Contains          // Match if the value contains the text
            };

            // Prepare the style with yellow fill
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            // Iterate through all cells that contain the text "Pending"
            Cell previousCell = null;
            Cell foundCell = cells.Find("Pending", previousCell, findOptions);
            while (foundCell != null)
            {
                // Apply the yellow fill style to the found cell
                foundCell.SetStyle(highlightStyle);

                // Continue searching from the next cell
                previousCell = foundCell;
                foundCell = cells.Find("Pending", previousCell, findOptions);
            }

            // Save the workbook with highlighted cells
            workbook.Save("HighlightedPending.xlsx");
        }
    }
}