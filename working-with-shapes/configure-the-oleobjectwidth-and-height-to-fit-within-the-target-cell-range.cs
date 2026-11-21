// Title: Setting OleObject dimensions to fit a target cell range with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells resize OLE object to cover cells B2:D5 | calculate total pixel width and height of a cell range for OleObject placement in Aspose.Cells | Aspose.Cells set OleObject placement to MoveAndSize programmatically | embed OLE file into a worksheet and fit it inside a specific cell block using Aspose.Cells | determine column width in pixels with Aspose.Cells for OLE object sizing
// Tags: resize OleObject based on cell range dimensions | calculate cell range pixel size Aspose.Cells | set OleObject placement MoveAndSize Aspose.Cells | embed OLE document into worksheet C# | use GetColumnWidthPixel for OLE sizing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, defines a cell range (B2:D5), reads an OLE file into a byte array, adds the OLE object to the sheet, computes the range's total pixel width and height using GetColumnWidthPixel and GetRowHeightPixel, assigns these values to the OleObject's Width and Height, sets its Placement to MoveAndSize, and saves the file as OleObjectFit.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range (e.g., B2:D5)
            int startRow = 1;      // B2 row (zero‑based)
            int startColumn = 1;   // B2 column (zero‑based)
            int endRow = 4;        // D5 row
            int endColumn = 3;     // D5 column

            // Path to the OLE source file
            string oleFilePath = "sample.docx";

            // Verify the OLE source file exists before adding
            if (File.Exists(oleFilePath))
            {
                // Read the file into a byte array (required by OleObjects.Add)
                byte[] oleData = File.ReadAllBytes(oleFilePath);

                // Add the OLE object; Add returns the index of the newly added object
                int oleIndex = sheet.OleObjects.Add(startRow, startColumn, 0, 0, oleData);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Calculate total width of the range in pixels
                double totalWidth = 0;
                for (int col = startColumn; col <= endColumn; col++)
                {
                    totalWidth += sheet.Cells.GetColumnWidthPixel(col);
                }

                // Calculate total height of the range in pixels
                double totalHeight = 0;
                for (int row = startRow; row <= endRow; row++)
                {
                    totalHeight += sheet.Cells.GetRowHeightPixel(row);
                }

                // Set the OLE object's size to match the cell range
                ole.Width = (int)totalWidth;
                ole.Height = (int)totalHeight;

                // Ensure the object moves and resizes with the cells
                ole.Placement = PlacementType.MoveAndSize;
            }
            else
            {
                Console.WriteLine($"Warning: OLE source file '{oleFilePath}' not found. Skipping OLE insertion.");
            }

            // Save the workbook
            string outputPath = "OleObjectFit.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
