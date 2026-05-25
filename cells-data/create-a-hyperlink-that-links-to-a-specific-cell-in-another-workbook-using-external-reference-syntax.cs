using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExternalReference
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the external reference address.
            // Format: [FileName]SheetName!CellAddress
            // Example links to cell B2 in Sheet1 of ExternalWorkbook.xlsx
            string externalAddress = "[ExternalWorkbook.xlsx]Sheet1!B2";

            // Add a hyperlink to cell A1 that points to the external cell.
            // Parameters: start cell name, rows, columns, address.
            sheet.Hyperlinks.Add("A1", 1, 1, externalAddress);

            // Optionally set the display text for the hyperlink
            sheet.Cells["A1"].PutValue("Go to External B2");

            // Save the workbook
            workbook.Save("HyperlinkToExternalCell.xlsx");
        }
    }
}