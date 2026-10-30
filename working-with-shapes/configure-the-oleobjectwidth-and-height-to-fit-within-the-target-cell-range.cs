// Title: Resize an OleObject to match a cell range size with Aspose.Cells for .NET
// Common Searches: how to programmatically set OleObject width and height based on a worksheet range in Aspose.Cells C# | Aspose.Cells calculate pixel width of B2:D5 range for OLE object sizing | C# resize embedded OLE object to fit specific cells using Aspose.Cells | adjust OleObject dimensions to match cell range dimensions in .NET workbook | determine total pixel size of a cell range for OleObject scaling with Aspose.Cells
// Tags: OleObject resize based on cell range Aspose.Cells | calculate range pixel dimensions C# | set OleObject width height programmatically | fit embedded OLE object within worksheet cells | Aspose.Cells adjust OLE object size .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// The example loads or creates a workbook, defines a target range (e.g., B2:D5), computes the total pixel width and height of that range, and then sets the first OleObject's Width and Height properties to those values before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Console.WriteLine($"Input file not found. Created a new workbook.");
            }

            Worksheet sheet = workbook.Worksheets[0];

            // Define the target cell range for the OLE object
            string targetRange = "B2:D5";
            AsposeRange range = sheet.Cells.CreateRange(targetRange);

            // Calculate total width of the range in pixels
            double totalWidth = 0;
            int firstCol = range.FirstColumn;
            int lastCol = range.FirstColumn + range.ColumnCount - 1;
            for (int col = firstCol; col <= lastCol; col++)
            {
                totalWidth += sheet.Cells.GetColumnWidthPixel(col);
            }

            // Calculate total height of the range in pixels
            double totalHeight = 0;
            int firstRow = range.FirstRow;
            int lastRow = range.FirstRow + range.RowCount - 1;
            for (int row = firstRow; row <= lastRow; row++)
            {
                totalHeight += sheet.Cells.GetRowHeightPixel(row);
            }

            // Adjust the first OLE object's size if it exists
            if (sheet.OleObjects.Count > 0)
            {
                OleObject ole = sheet.OleObjects[0];
                ole.Width = (int)totalWidth;   // OleObject.Width expects int
                ole.Height = (int)totalHeight; // OleObject.Height expects int
            }
            else
            {
                Console.WriteLine("No OLE objects found on the worksheet.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
