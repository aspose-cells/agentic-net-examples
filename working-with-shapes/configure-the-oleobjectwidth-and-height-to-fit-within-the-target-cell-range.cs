// Title: Resize an OleObject to exactly fit the B2:D5 cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Calculate the pixel width and height of a specified Excel cell range and assign those values to OleObject.Width and OleObject.Height so the embedded Word document fills the range. | Add a Word document as an OLE object at cell B2 and programmatically adjust its size to match the combined column widths and row heights of cells B2 through D5 with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# set OleObject dimensions to match a cell range | how to compute pixel size from Excel column width and row height for OLE objects | fit embedded Word document to specific cells using Aspose.Cells .NET | resize OleObject to cover B2:D5 range programmatically | convert Excel column width to pixels Aspose.Cells example
// Tags: oleobject size adjustment based on cell range | pixel dimension calculation from column width Aspose.Cells | embed word document as oleobject in worksheet c# | set oleobject width and height programmatically | fit oleobject to specific excel range asp.net

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example creates a workbook, computes the total column width and row height of the B2:D5 range in pixels, loads a Word file, adds it as an OLE object at the start cell, sets OleObject.Width and OleObject.Height to the calculated pixel values, and saves the worksheet as output.xlsx.
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
                string startCellName = "B2";
                string endCellName = "D5";

                Cell startCell = sheet.Cells[startCellName];
                Cell endCell = sheet.Cells[endCellName];

                int startRow = startCell.Row;
                int endRow = endCell.Row;
                int startCol = startCell.Column;
                int endCol = endCell.Column;

                // Calculate total column width in characters for the range
                double totalColumnWidth = 0;
                for (int col = startCol; col <= endCol; col++)
                {
                    totalColumnWidth += sheet.Cells.GetColumnWidth(col);
                }
                // Approximate conversion: 1 character width ≈ 7 pixels
                int widthInPixels = (int)Math.Round(totalColumnWidth * 7);

                // Calculate total row height in points for the range
                double totalRowHeight = 0;
                for (int row = startRow; row <= endRow; row++)
                {
                    totalRowHeight += sheet.Cells.GetRowHeight(row);
                }
                // Approximate conversion: 1 point ≈ 1.33 pixels
                int heightInPixels = (int)Math.Round(totalRowHeight * 1.33);

                // Load OLE object data (sample.docx) if the file exists
                string oleFilePath = "sample.docx";
                byte[] oleData;
                if (File.Exists(oleFilePath))
                {
                    oleData = File.ReadAllBytes(oleFilePath);
                }
                else
                {
                    Console.WriteLine($"File '{oleFilePath}' not found. OLE object will be empty.");
                    oleData = new byte[0];
                }

                // Add an OLE object (Word document) to the worksheet
                // Note: overload expects data first, then the ProgID
                int oleIndex = sheet.OleObjects.Add(startRow, startCol, heightInPixels, widthInPixels, oleData, "Word.Document");

                // Configure the OLE object's size to fit the target cell range
                OleObject ole = sheet.OleObjects[oleIndex];
                ole.Width = widthInPixels;
                ole.Height = heightInPixels;

                // Save the workbook
                workbook.Save("output.xlsx");
                Console.WriteLine("Workbook saved as output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
