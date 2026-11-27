// Title: How to batch‑insert a SVG company logo into the top‑right corner of every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add the same SVG image to all worksheets in a workbook | place a free‑floating picture at the top right of each Excel sheet with Aspose.Cells | C# code to insert a logo SVG after the last used column in every worksheet | batch add company logo to multiple worksheets using Aspose.Cells .NET | set picture placement to FreeFloating when adding SVG to Excel with Aspose.Cells
// Tags: batch insert SVG picture Aspose.Cells | add floating image to Excel worksheets C# | insert picture at column after data Aspose.Cells | logo placement in worksheet corner Aspose.Cells | FreeFloating picture mode Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example validates the source XLSX workbook and SVG logo file, loads the workbook, and iterates through each worksheet. For every sheet it calculates the column immediately after the last used column, adds the SVG image (100 × 100 px) at row 0 using Pictures.Add, sets the picture's Placement to FreeFloating so it stays independent of cells, and finally saves the modified workbook as a new XLSX file.
class Program
{
    static void Main()
    {
        // Paths – adjust these to your environment
        string workbookPath = "input.xlsx";   // source workbook
        string logoSvgPath = "company_logo.svg"; // SVG logo file
        string outputPath = "output_with_logo.xlsx"; // result workbook

        // Validate input files
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        if (!File.Exists(logoSvgPath))
        {
            Console.WriteLine($"Error: SVG logo file not found at '{logoSvgPath}'.");
            return;
        }

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Desired picture size in pixels (you can change these values)
        const int pictureHeight = 100;
        const int pictureWidth = 100;

        // Iterate over every worksheet and insert the SVG logo at the top‑right corner
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the column after the last used column to place the logo
            int lastUsedColumn = sheet.Cells.MaxColumn; // zero‑based index of the rightmost used column
            int targetColumn = lastUsedColumn + 1;      // one column to the right of existing data
            int targetRow = 0;                          // first row (top)

            // Add the SVG picture. Aspose.Cells supports SVG as a picture format.
            int pictureIndex = sheet.Pictures.Add(targetRow, targetColumn, pictureHeight, pictureWidth, logoSvgPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Optional: make the picture free‑floating so it doesn't move with cells
            picture.Placement = PlacementType.FreeFloating;
        }

        // Save the modified workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Successfully inserted the logo into all worksheets. Saved as '{outputPath}'.");
    }
}
