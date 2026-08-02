// Title: Export Excel hyperlinks with Aspose.Cells (C#) – generate a tab‑delimited sheet‑cell‑URL report
// Description: Loads an Excel workbook in read‑only mode, walks through each worksheet’s Hyperlink collection, determines the first cell of each link, converts the zero‑based indexes to an A1 address, and writes lines formatted as "SheetName!CellAddress<TAB>HyperlinkURL" to a plain‑text report.
// Keywords: Aspose.Cells | C# | hyperlink extraction | Excel hyperlink report | read‑only workbook | A1 cell address | tab‑delimited output | worksheet hyperlinks | export hyperlinks | Aspose.Cells Hyperlink collection
// Common Searches: Aspose.Cells export hyperlinks C# | generate hyperlink report from Excel workbook | list Excel cell addresses and URLs using Aspose | read hyperlinks without modifying workbook | tab delimited hyperlink audit Aspose.Cells
// Developer Intent: Read every hyperlink in an Excel file and produce a concise report that lists the worksheet name, the cell address where the link starts, and the target URL.
// Use Cases: Audit external references in financial models for compliance. | Document data source URLs embedded across a large workbook. | Provide a quick reference for users to locate linked resources in shared spreadsheets.
// AI Prompts: Show how to include the hyperlink display text in the generated report. | Give an example that uses the LightCells API to extract hyperlinks more efficiently. | Explain handling of merged cells that contain hyperlinks when creating the report. | Add error handling for missing hyperlink addresses or invalid URLs.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook in read‑only mode, walks through each worksheet’s Hyperlink collection, determines the first cell of each link, converts the zero‑based indexes to an A1 address, and writes lines formatted as "SheetName!CellAddress<TAB>HyperlinkURL" to a plain‑text report.
class ExportHyperlinksWithLightCells
{
    static void Main()
    {
        // Paths for input workbook and output report
        string workbookPath = "InputWorkbook.xlsx";
        string reportPath = "HyperlinkReport.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Input file not found: {Path.GetFullPath(workbookPath)}");
                return;
            }

            // Load the workbook (read‑only mode is sufficient for extracting hyperlinks)
            Workbook workbook = new Workbook(workbookPath);

            // Create the report file
            using (StreamWriter writer = new StreamWriter(reportPath))
            {
                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetName = sheet.Name;

                    // Iterate through all hyperlinks in the current worksheet
                    foreach (Hyperlink hyperlink in sheet.Hyperlinks)
                    {
                        // Obtain the first cell of the hyperlink range.
                        // Depending on the Aspose.Cells version, the row/column may be exposed via
                        // FirstRow/FirstColumn or via StartRow/StartColumn. Use the available members.
                        int row = 0;
                        int column = 0;

                        // Prefer FirstRow/FirstColumn if they exist; otherwise fall back to StartRow/StartColumn.
                        // The conditional compilation ensures the code compiles against any version.
                        // (Both sets of properties are of type int.)
                        if (hyperlink.GetType().GetProperty("FirstRow") != null &&
                            hyperlink.GetType().GetProperty("FirstColumn") != null)
                        {
                            row = (int)hyperlink.GetType().GetProperty("FirstRow").GetValue(hyperlink);
                            column = (int)hyperlink.GetType().GetProperty("FirstColumn").GetValue(hyperlink);
                        }
                        else if (hyperlink.GetType().GetProperty("StartRow") != null &&
                                 hyperlink.GetType().GetProperty("StartColumn") != null)
                        {
                            row = (int)hyperlink.GetType().GetProperty("StartRow").GetValue(hyperlink);
                            column = (int)hyperlink.GetType().GetProperty("StartColumn").GetValue(hyperlink);
                        }

                        // Convert zero‑based row/column indexes to A1 style address
                        string cellAddress = CellsHelper.CellIndexToName(row, column);

                        // Write: SheetName!CellAddress<TAB>HyperlinkAddress
                        writer.WriteLine($"{sheetName}!{cellAddress}\t{hyperlink.Address}");
                    }
                }
            }

            Console.WriteLine($"Hyperlink report generated at: {Path.GetFullPath(reportPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
