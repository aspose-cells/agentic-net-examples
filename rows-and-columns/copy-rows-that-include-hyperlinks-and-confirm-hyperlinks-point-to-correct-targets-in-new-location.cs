// Title: Copy a worksheet row containing a hyperlink and verify the hyperlink addresses with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells CopyRows with CopyOptions.ExtendToAdjacentRange to copy row 2 (A2) to row 4, then iterate the Hyperlinks collection to output each hyperlink’s cell address and URL. | Create a workbook, add a hyperlink to cell A2, duplicate the entire row to a new location while preserving the link, and confirm that the copied hyperlink still points to the original target URL.
// Common Searches: Aspose.Cells copy row with hyperlink and keep link address | How to preserve hyperlink when using CopyRows in C# | ExtendToAdjacentRange option effect on hyperlink ranges Aspose.Cells | Retrieve hyperlink addresses after copying rows in a worksheet | Validate copied hyperlink target URL with Aspose.Cells .NET
// Tags: CopyRows hyperlink preservation Aspose.Cells | ExtendToAdjacentRange copy option | enumerate worksheet hyperlinks C# | validate hyperlink target after row copy | Aspose.Cells row duplication with links

using System;
using Aspose.Cells;

namespace HyperlinkRowCopyDemo
{
    // The example creates a workbook, inserts a hyperlink into cell A2, then copies row 2 to row 4 using CopyOptions with ExtendToAdjacentRange set to true. After the copy, it enumerates the worksheet's Hyperlinks collection to display each hyperlink’s cell address and destination URL, confirming that the hyperlink was preserved, and finally saves the workbook as HyperlinkRowCopyResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data in row 1 (index 0) and a hyperlink in row 2 (index 1)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Link Cell");
            // Hyperlink points to an external website
            sheet.Hyperlinks.Add("A2", 1, 1, "https://www.aspose.com");

            // Prepare copy options: ExtendToAdjacentRange = true
            // This ensures the hyperlink range is extended rather than creating a new hyperlink object
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ExtendToAdjacentRange = true;

            // Copy row 2 (index 1) to row 4 (index 3)
            // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber, copyOptions
            sheet.Cells.CopyRows(sheet.Cells, 1, 3, 1, copyOptions);

            // Verify hyperlinks after copy
            Console.WriteLine("Total hyperlinks in worksheet: " + sheet.Hyperlinks.Count);

            // Find hyperlink(s) in the original row (A2) and the copied row (A4)
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                // Determine the cell address of the hyperlink using its Area property
                string cellAddress = new StyleFlag().ToString(); // placeholder not needed
                string startCell = CellIndexToName(link.Area.StartRow, link.Area.StartColumn);
                Console.WriteLine($"Hyperlink at {startCell} points to: {link.Address}");
            }

            // Save the workbook (optional, demonstrates lifecycle rule)
            workbook.Save("HyperlinkRowCopyResult.xlsx");
        }

        // Helper method to convert zero‑based row/column indexes to Excel cell name (e.g., 0,0 -> A1)
        private static string CellIndexToName(int row, int column)
        {
            // Convert column index to letters
            string columnName = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            // Row index is zero‑based, so add 1
            return $"{columnName}{row + 1}";
        }
    }
}
