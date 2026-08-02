using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkDemo
{
    class Program
    {
        static void Main()
        {
            // Array of file paths to be added to the worksheet
            string[] filePaths = new string[]
            {
                @"C:\Docs\Report1.pdf",
                @"C:\Docs\Report2.pdf",
                @"C:\Docs\Report3.pdf"
            };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Import the file paths vertically starting at cell A1 (row 0, column 0)
            // Using the ImportArray method with isVertical = true
            worksheet.Cells.ImportArray(filePaths, 0, 0, true);

            // Add a hyperlink in the adjacent column (B) for each file path
            for (int i = 0; i < filePaths.Length; i++)
            {
                // Set display text for the hyperlink cell (optional)
                worksheet.Cells[i, 1].PutValue("Open");

                // Add hyperlink to cell B(i+1) pointing to the corresponding file path
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, address
                worksheet.Hyperlinks.Add(i, 1, 1, 1, filePaths[i]);
            }

            // Save the workbook to an Excel file
            workbook.Save("FilePathsWithHyperlinks.xlsx");
        }
    }
}