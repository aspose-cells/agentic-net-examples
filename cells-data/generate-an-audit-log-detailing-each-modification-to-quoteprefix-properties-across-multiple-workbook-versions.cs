using System;
using Aspose.Cells;
using Aspose.Cells.Revisions;

class QuotePrefixAudit
{
    static void Main()
    {
        // ---------- Create initial workbook ----------
        Workbook wb = new Workbook();
        wb.Settings.Shared = true;                     // enable revision tracking
        wb.Settings.QuotePrefixToStyle = true;         // apply QuotePrefix style for leading '
        wb.Worksheets[0].Cells["A1"].PutValue("'First"); // cell with leading quote
        string filePath = "QuotePrefixAudit.xlsx";
        wb.Save(filePath, SaveFormat.Xlsx);            // save version 1

        // ---------- Load workbook and modify ----------
        Workbook wb2 = new Workbook(filePath);
        wb2.Settings.QuotePrefixToStyle = false;       // disable automatic QuotePrefix styling
        wb2.Worksheets[0].Cells["A1"].PutValue("'Second"); // modify same cell
        wb2.Save(filePath, SaveFormat.Xlsx);           // save version 2 (overwrites)

        // ---------- Load final workbook and audit ----------
        Workbook finalWb = new Workbook(filePath);
        if (!finalWb.HasRevisions)
        {
            Console.WriteLine("No revisions found.");
            return;
        }

        // Iterate through all revision logs
        foreach (RevisionLog log in finalWb.Worksheets.RevisionLogs)
        {
            foreach (Revision rev in log.Revisions)
            {
                // We're interested only in cell changes
                if (rev is RevisionCellChange cellChange)
                {
                    bool oldQuote = cellChange.OldStyle?.QuotePrefix ?? false;
                    bool newQuote = cellChange.NewStyle?.QuotePrefix ?? false;

                    // Log only when QuotePrefix property changed
                    if (oldQuote != newQuote)
                    {
                        Console.WriteLine($"Cell {cellChange.CellName} QuotePrefix changed from {oldQuote} to {newQuote} (Revision ID: {cellChange.Id})");
                    }
                }
            }
        }
    }
}