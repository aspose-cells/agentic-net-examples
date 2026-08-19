// Title: Audit QuotePrefix changes across workbook versions with Aspose.Cells for .NET
// Description: C# program that creates two workbook revisions, toggles QuotePrefixToStyle, inserts values with leading apostrophes, extracts the QuotePrefix flag from every used cell, compares the snapshots, and outputs a console audit log of cells whose QuotePrefix state changed.
// Keywords: Aspose.Cells | .NET | QuotePrefix | QuotePrefixToStyle | audit log | workbook version comparison | cell style tracking | Excel revision audit | C# | Aspose.Cells API
// Common Searches: Aspose.Cells how to track QuotePrefix changes | C# audit QuotePrefix property between two workbooks | compare cell style QuotePrefix with Aspose.Cells | log QuotePrefix modifications in Excel using Aspose | extract QuotePrefix flag from cells Aspose.Cells
// Developer Intent: Generate a detailed audit log that identifies cells whose QuotePrefix flag changed between different workbook revisions using Aspose.Cells for .NET.
// Use Cases: Validate that enabling QuotePrefixToStyle correctly applies the QuotePrefix style to cells with leading apostrophes. | Detect removal of QuotePrefix style after disabling automatic handling. | Create compliance or change‑management reports showing QuotePrefix usage across spreadsheet revisions. | Integrate the audit into CI pipelines to ensure style consistency across generated workbooks.
// AI Prompts: Add an option to export the audit differences to a CSV or JSON file. | Extend the program to iterate over all worksheets and include worksheet names in the log. | Write unit tests using NUnit or xUnit that confirm QuotePrefix changes are detected for various scenarios. | Implement a command‑line interface that accepts file paths and version identifiers. | Provide a PowerShell script that runs the audit on a folder of workbook revisions.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Revisions;

namespace QuotePrefixAudit
{
    // C# program that creates two workbook revisions, toggles QuotePrefixToStyle, inserts values with leading apostrophes, extracts the QuotePrefix flag from every used cell, compares the snapshots, and outputs a console audit log of cells whose QuotePrefix state changed.
    class Program
    {
        // Simple data structure to hold QuotePrefix info for a cell
        class CellQuoteInfo
        {
            public string Address { get; set; }
            public bool QuotePrefix { get; set; }
        }

        // Extract QuotePrefix information from all used cells in a worksheet
        static List<CellQuoteInfo> GetQuotePrefixInfo(Worksheet sheet)
        {
            var list = new List<CellQuoteInfo>();
            // Determine the used range to limit iteration
            var maxRow = sheet.Cells.MaxDataRow;
            var maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];
                    // Skip empty cells
                    if (cell == null || cell.Type == CellValueType.IsNull) continue;

                    var style = cell.GetStyle();
                    list.Add(new CellQuoteInfo
                    {
                        Address = cell.Name,
                        QuotePrefix = style.QuotePrefix
                    });
                }
            }
            return list;
        }

        // Compare two snapshots and output differences
        static void AuditDifferences(string versionOld, List<CellQuoteInfo> oldInfo,
                                     string versionNew, List<CellQuoteInfo> newInfo)
        {
            // Build dictionaries for quick lookup
            var oldDict = new Dictionary<string, bool>();
            foreach (var info in oldInfo) oldDict[info.Address] = info.QuotePrefix;

            var newDict = new Dictionary<string, bool>();
            foreach (var info in newInfo) newDict[info.Address] = info.QuotePrefix;

            // Union of all addresses
            var allAddresses = new HashSet<string>(oldDict.Keys);
            allAddresses.UnionWith(newDict.Keys);

            foreach (var addr in allAddresses)
            {
                bool oldVal = oldDict.ContainsKey(addr) ? oldDict[addr] : false;
                bool newVal = newDict.ContainsKey(addr) ? newDict[addr] : false;

                if (oldVal != newVal)
                {
                    Console.WriteLine($"Cell {addr}: QuotePrefix changed from {oldVal} (v{versionOld}) to {newVal} (v{versionNew})");
                }
            }
        }

        static void Main()
        {
            // ---------- Create first version ----------
            Workbook wbV1 = new Workbook();                         // create workbook
            wbV1.Settings.QuotePrefixToStyle = true;               // enable automatic QuotePrefix style
            Cell cellA1 = wbV1.Worksheets[0].Cells["A1"];
            cellA1.PutValue("'Alpha");                             // leading apostrophe triggers QuotePrefix
            wbV1.Save("Workbook_V1.xlsx");                         // save first version

            // ---------- Create second version ----------
            // Load the first version to continue modifications
            Workbook wbV2 = new Workbook("Workbook_V1.xlsx");       // load workbook
            wbV2.Settings.QuotePrefixToStyle = false;              // disable automatic QuotePrefix style
            Cell cellA2 = wbV2.Worksheets[0].Cells["A2"];
            cellA2.PutValue("'Beta");                              // value stored literally, no QuotePrefix
            wbV2.Save("Workbook_V2.xlsx");                         // save second version

            // ---------- Load versions for auditing ----------
            Workbook loadV1 = new Workbook("Workbook_V1.xlsx");     // load first version
            Workbook loadV2 = new Workbook("Workbook_V2.xlsx");     // load second version

            // Extract QuotePrefix info from each version
            var infoV1 = GetQuotePrefixInfo(loadV1.Worksheets[0]);
            var infoV2 = GetQuotePrefixInfo(loadV2.Worksheets[0]);

            // ---------- Generate audit log ----------
            Console.WriteLine("Audit Log of QuotePrefix modifications:");
            AuditDifferences("1", infoV1, "2", infoV2);

            // Keep console window open when run outside IDE
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
