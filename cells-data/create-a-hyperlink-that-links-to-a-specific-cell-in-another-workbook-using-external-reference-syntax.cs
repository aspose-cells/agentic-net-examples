// Title: Add an Excel external‑workbook hyperlink to a specific cell with Aspose.Cells (C#)
// Description: Demonstrates how to construct an Excel external reference like "'[ExternalWorkbook.xlsx]Sheet1'!B2", add it as a hyperlink to cell A1, set custom display text, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel external hyperlink | external workbook reference | cell link | hyperlink.Add | Excel reference syntax | cross‑workbook navigation | GitHub example
// Common Searches: Aspose.Cells create hyperlink to another workbook | C# external reference format for Excel hyperlink | How to link to a cell in a different Excel file using Aspose.Cells | Add external workbook hyperlink with Aspose.Cells C# | Excel cross‑workbook hyperlink code sample
// Developer Intent: Insert a hyperlink that opens a designated cell in a separate Excel workbook.
// Use Cases: Build an index workbook that jumps to detailed records stored in separate files. | Automate cross‑file navigation links for large financial models. | Generate summary reports that point to source data cells in external workbooks.
// AI Prompts: Write C# Aspose.Cells code to add a hyperlink in cell B3 that opens cell D10 in "Metrics.xlsx" sheet "Summary". | Show how to create the external reference string for a hyperlink to cell A1 in "SourceData.xlsx" sheet "Data" using Aspose.Cells. | Explain how to set custom display text for an external workbook hyperlink and save the file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExternalReference
{
    // Demonstrates how to construct an Excel external reference like "'[ExternalWorkbook.xlsx]Sheet1'!B2", add it as a hyperlink to cell A1, set custom display text, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the external workbook file name, sheet name and target cell
            string externalFileName = "ExternalWorkbook.xlsx";
            string externalSheetName = "Sheet1";
            string externalCell = "B2";

            // Build the external reference address in the format required by Excel
            // Example: '[ExternalWorkbook.xlsx]Sheet1'!B2
            string address = $"'[${externalFileName}]${externalSheetName}'!{externalCell}";

            // Add a hyperlink to cell A1 that points to the external cell
            // Parameters: start cell name, total rows, total columns, address
            worksheet.Hyperlinks.Add("A1", 1, 1, address);

            // Optionally set the display text for the hyperlink
            worksheet.Cells["A1"].PutValue("Go to External Cell");

            // Save the workbook
            workbook.Save("HyperlinkToExternalCell.xlsx");
        }
    }
}
