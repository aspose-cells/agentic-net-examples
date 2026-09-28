// Title: Add a hyperlink in Aspose.Cells for .NET that points to a specific cell in another workbook using Excel external reference syntax
// AI Prompts: Generate C# code with Aspose.Cells that creates a hyperlink in cell A1 of a workbook, linking to cell A1 of Sheet1 in an external Excel file. | Show how to register an external workbook in Aspose.Cells and build the "'[File]Sheet'!Cell" reference for a hyperlink. | Provide a complete example that saves the main workbook after adding the external cell hyperlink.
// Common Searches: aspnet add hyperlink to cell in another Excel workbook using Aspose.Cells | how to use external link entry with Aspose.Cells C# | Aspose.Cells external reference syntax for hyperlink to external workbook cell | create hyperlink that opens a specific cell in a different workbook with Aspose.Cells .NET
// Tags: add external hyperlink Aspose.Cells | external link entry Aspose.Cells C# | Excel external reference syntax Aspose.Cells | hyperlink to specific cell external workbook | Aspose.Cells workbook hyperlink external cell

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExternalCell
{
    // The example creates a main workbook, registers an external link for "ExternalWorkbook.xlsx", constructs an external reference address "'[ExternalWorkbook.xlsx]Sheet1'!A1", adds a hyperlink in cell A1 that points to that external cell, sets display text, and saves the file as "HyperlinkToExternalCell.xlsx".
    class Program
    {
        static void Main()
        {
            // Create the main workbook
            Workbook mainWb = new Workbook();
            Worksheet sheet = mainWb.Worksheets[0];

            // Add an external link entry for the target workbook (required for external references)
            // The external workbook file name and the sheet(s) that will be referenced
            string externalFile = "ExternalWorkbook.xlsx";
            string[] externalSheets = new string[] { "Sheet1" };
            int linkIndex = sheet.Workbook.Worksheets.ExternalLinks.Add(externalFile, externalSheets);
            ExternalLink extLink = sheet.Workbook.Worksheets.ExternalLinks[linkIndex];

            // Build the external reference address that points to cell A1 in Sheet1 of the external workbook
            // Excel external reference syntax: '[FileName]SheetName'!CellAddress
            string address = $"'[${externalFile}]Sheet1'!A1";

            // Add a hyperlink to cell A1 of the main workbook that points to the external cell
            // Using the overload that takes a cell name, row/column span and the address
            sheet.Hyperlinks.Add("A1", 1, 1, address);

            // Optionally set display text for the hyperlink
            sheet.Cells["A1"].PutValue("Go to External A1");

            // Save the workbook
            mainWb.Save("HyperlinkToExternalCell.xlsx");
        }
    }
}
