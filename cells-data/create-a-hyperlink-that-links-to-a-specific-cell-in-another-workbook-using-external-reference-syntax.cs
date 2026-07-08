using System;
using Aspose.Cells;

class HyperlinkExternalCellDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the external workbook file name, sheet and cell to link to
        string externalFileName = "ExternalWorkbook.xlsx";
        string externalSheetName = "Sheet1";
        string externalCellAddress = "B2";

        // Build the external reference address in the format:
        // '[ExternalWorkbook.xlsx]Sheet1'!B2
        string hyperlinkAddress = $"'[${externalFileName}]${externalSheetName}'!{externalCellAddress}";

        // Add a hyperlink to cell A1 that points to the external cell
        // Parameters: start cell name, rows, columns, address
        worksheet.Hyperlinks.Add("A1", 1, 1, hyperlinkAddress);

        // Set the display text for the hyperlink
        worksheet.Cells["A1"].PutValue("Link to external cell");

        // Save the workbook
        workbook.Save("HyperlinkExternalCellDemo.xlsx");
    }
}