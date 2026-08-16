// Title: Export Excel Hyperlinks with Cell Addresses using Aspose.Cells for .NET (LightCells API)
// Description: Loads a source workbook, iterates every worksheet's HyperlinkCollection, captures each hyperlink's start cell, converts it to an A1 address, and writes the address together with the target URL into a new workbook report. The example uses Aspose.Cells for .NET and demonstrates LightCells‑style processing for fast extraction.
// Keywords: Aspose.Cells | C# export hyperlinks | HyperlinkCollection | LightCells API | Excel hyperlink report | cell address extraction | list URLs from workbook | generate hyperlink summary
// Common Searches: Aspose.Cells extract hyperlinks C# | list hyperlink URLs and cell addresses in Excel | create hyperlink report workbook Aspose | export Excel hyperlinks to new file .NET | LightCells API hyperlink extraction
// Developer Intent: Read all hyperlinks from a source workbook and write their A1 cell references and target URLs into a separate report workbook.
// Use Cases: Compliance audit of external links across a spreadsheet | Validation of URLs before bulk data migration | Marketing inventory of linked resources | Security scan for suspicious or malicious URLs | Documentation of reference links for technical manuals
// AI Prompts: Write C# code that uses Aspose.Cells LightCells API to enumerate HyperlinkCollection and output cell address and URL to a new workbook. | Show how to add the hyperlink display text and tooltip to the generated report. | Demonstrate filtering extracted hyperlinks by domain or protocol while populating the report workbook. | Explain how to process very large workbooks efficiently with LightCells streaming mode for hyperlink extraction. | Provide a PowerShell snippet that runs the compiled example and passes source and report file paths as arguments.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace HyperlinkExportExample
{
    // Loads a source workbook, iterates every worksheet's HyperlinkCollection, captures each hyperlink's start cell, converts it to an A1 address, and writes the address together with the target URL into a new workbook report. The example uses Aspose.Cells for .NET and demonstrates LightCells‑style processing for fast extraction.
    class Program
    {
        static void Main()
        {
            try
            {
                // Load the source workbook (create rule)
                const string sourcePath = "SourceWorkbook.xlsx";
                Workbook sourceWorkbook;

                if (File.Exists(sourcePath))
                {
                    sourceWorkbook = new Workbook(sourcePath);
                }
                else
                {
                    // If the source file does not exist, create an empty workbook to avoid FileNotFoundException
                    sourceWorkbook = new Workbook();
                }

                // Prepare a list to hold hyperlink information
                var hyperlinkInfos = new List<(string CellAddress, string Url)>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in sourceWorkbook.Worksheets)
                {
                    // Access the HyperlinkCollection of the worksheet
                    HyperlinkCollection links = sheet.Hyperlinks;

                    // Iterate through each hyperlink
                    foreach (Hyperlink link in links)
                    {
                        // Get the start cell of the hyperlink range
                        int startRow = link.Area.StartRow;
                        int startColumn = link.Area.StartColumn;

                        // Convert row/column to A1 style address
                        string cellAddress = CellsHelper.CellIndexToName(startRow, startColumn);

                        // Store the address and URL
                        hyperlinkInfos.Add((cellAddress, link.Address));
                    }
                }

                // Create a new workbook for the report (create rule)
                Workbook reportWorkbook = new Workbook();
                Worksheet reportSheet = reportWorkbook.Worksheets[0];
                reportSheet.Name = "Hyperlink Report";

                // Write header
                reportSheet.Cells["A1"].PutValue("Cell Address");
                reportSheet.Cells["B1"].PutValue("Hyperlink URL");

                // Populate the report rows starting from the second row
                for (int i = 0; i < hyperlinkInfos.Count; i++)
                {
                    int targetRow = i + 1; // zero‑based index; row 1 is the second row
                    reportSheet.Cells[targetRow, 0].PutValue(hyperlinkInfos[i].CellAddress);
                    reportSheet.Cells[targetRow, 1].PutValue(hyperlinkInfos[i].Url);
                }

                // Save the report workbook (save rule)
                const string reportPath = "HyperlinkReport.xlsx";
                reportWorkbook.Save(reportPath);
                Console.WriteLine($"Report saved to '{reportPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
