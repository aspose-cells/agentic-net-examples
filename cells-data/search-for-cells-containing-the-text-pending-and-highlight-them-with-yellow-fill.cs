using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightPending
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure find options to search for cells whose values contain the text "Pending"
            FindOptions findOptions = new FindOptions
            {
                LookInType = LookInType.Values,      // Search in cell values
                LookAtType = LookAtType.Contains      // Match if the value contains the target text
            };

            // Repeatedly find cells containing "Pending"
            Cell previousCell = null;
            while (true)
            {
                Cell foundCell = worksheet.Cells.Find("Pending", previousCell, findOptions);
                if (foundCell == null)
                    break; // No more matches

                // Create a style with yellow fill
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.Yellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Apply the style to the found cell
                foundCell.SetStyle(highlightStyle);

                // Continue searching from the cell just found
                previousCell = foundCell;
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}