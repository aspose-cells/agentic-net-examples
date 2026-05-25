using System;
using Aspose.Cells;
using Aspose.Cells.Revisions;

class QuotePrefixAudit
{
    static void Main()
    {
        // Create a shared workbook and enable QuotePrefixToStyle
        Workbook wb = new Workbook();
        wb.Settings.Shared = true;
        wb.Settings.QuotePrefixToStyle = true;

        // Insert a value that starts with an apostrophe
        wb.Worksheets[0].Cells["A1"].PutValue("'First");
        // Save the first version
        wb.Save("QuotePrefix_v1.xlsx");

        // Load the first version, change the setting, and modify another cell
        Workbook wb2 = new Workbook("QuotePrefix_v1.xlsx");
        wb2.Settings.QuotePrefixToStyle = false;
        wb2.Worksheets[0].Cells["A2"].PutValue("'Second");
        // Save the second version
        wb2.Save("QuotePrefix_v2.xlsx");

        // Load the final workbook and generate an audit log of QuotePrefix changes
        Workbook finalWb = new Workbook("QuotePrefix_v2.xlsx");

        if (!finalWb.HasRevisions)
        {
            Console.WriteLine("No revisions found in the workbook.");
            return;
        }

        Console.WriteLine("Audit Log of QuotePrefix modifications:");
        foreach (RevisionLog log in finalWb.Worksheets.RevisionLogs)
        {
            foreach (Revision rev in log.Revisions)
            {
                if (rev is RevisionCellChange cellChange)
                {
                    bool oldQuotePrefix = cellChange.OldStyle?.QuotePrefix ?? false;
                    bool newQuotePrefix = cellChange.NewStyle?.QuotePrefix ?? false;

                    if (oldQuotePrefix != newQuotePrefix)
                    {
                        Console.WriteLine($"Cell {cellChange.CellName}: QuotePrefix changed from {oldQuotePrefix} to {newQuotePrefix}");
                        Console.WriteLine($"  Revision ID: {cellChange.Id}");
                        Console.WriteLine($"  Worksheet: {cellChange.Worksheet.Name}");
                    }
                }
            }
        }
    }
}