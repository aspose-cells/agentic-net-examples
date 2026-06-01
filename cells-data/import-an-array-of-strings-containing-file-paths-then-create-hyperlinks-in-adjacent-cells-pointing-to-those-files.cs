using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    class Program
    {
        static void Main()
        {
            // Array of file paths to be imported
            string[] filePaths = new string[]
            {
                @"C:\Documents\Report1.pdf",
                @"C:\Documents\Report2.pdf",
                @"C:\Documents\Report3.pdf"
            };

            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define where to place the file paths (column A) and hyperlinks (column B)
            int startRow = 0;          // Row index 0 => first row (A1)
            int pathColumn = 0;        // Column index 0 => column A
            int hyperlinkColumn = 1;   // Column index 1 => column B

            // Import the file path strings vertically into column A (ImportArray rule)
            worksheet.Cells.ImportArray(filePaths, startRow, pathColumn, true);

            // Add a hyperlink in the adjacent cell (column B) for each file path
            for (int i = 0; i < filePaths.Length; i++)
            {
                int currentRow = startRow + i;

                // Add hyperlink to the cell at (currentRow, hyperlinkColumn) (HyperlinkCollection.Add rule)
                worksheet.Hyperlinks.Add(currentRow, hyperlinkColumn, 1, 1, filePaths[i]);

                // Optional: set display text for the hyperlink
                worksheet.Cells[currentRow, hyperlinkColumn].PutValue("Open File");
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("FilePathsWithHyperlinks.xlsx");
        }
    }
}