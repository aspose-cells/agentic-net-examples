using System;
using Aspose.Cells;

namespace AsposeCellsHiddenExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and add a hidden external link.
            // ------------------------------------------------------------
            // The external link is added implicitly when a formula references
            // a cell in another workbook. This link is not visible as a
            // hyperlink in the UI, therefore it is considered "hidden".
            Workbook wb = new Workbook();

            // Add some data to the current sheet (optional, just for context)
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Local Value");
            ws.Cells["B1"].PutValue("External Link Demo");

            // Add a formula that points to an external workbook.
            // The external workbook does not need to exist at this moment;
            // Aspose.Cells will create the ExternalLink entry based on the formula.
            ws.Cells["C1"].Formula = "='[ExternalData.xlsx]Sheet1'!A1";

            // Save the workbook that now contains a hidden external link.
            wb.Save("HiddenExternalLink.xlsx");

            // ------------------------------------------------------------
            // 2. Load the workbook and inspect the hidden external link.
            // ------------------------------------------------------------
            Workbook loadedWb = new Workbook("HiddenExternalLink.xlsx");

            // The ExternalLinks collection holds all external references.
            ExternalLinkCollection externalLinks = loadedWb.Worksheets.ExternalLinks;

            // Iterate through the collection and display properties.
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // IsVisible is false for links that are not represented as
                // visible hyperlinks in Excel (i.e., hidden external links).
                Console.WriteLine($"External Link #{i}");
                Console.WriteLine($"  DataSource : {link.DataSource}");
                Console.WriteLine($"  IsVisible  : {link.IsVisible}");
                Console.WriteLine($"  IsReferred : {link.IsReferred}");
                Console.WriteLine($"  Type       : {link.Type}");
                Console.WriteLine();
            }

            // ------------------------------------------------------------
            // 3. Document potential use cases for hidden external hyperlinks.
            // ------------------------------------------------------------
            // Use Case 1: Data Consolidation
            //   - Pull values from multiple source workbooks without exposing
            //     the source paths to end‑users. Formulas reference external data
            //     while the workbook UI shows only the calculated results.

            // Use Case 2: Dynamic Dashboard Refresh
            //   - A dashboard workbook can reference a "master" data workbook.
            //     When the master file is updated, the dashboard automatically
            //     reflects the changes without revealing the link to users.

            // Use Case 3: Auditing & Traceability
            //   - Keep a hidden audit trail of source files used for calculations.
            //     Auditors can load the workbook programmatically and inspect
            //     ExternalLink.DataSource to verify data provenance.

            // Use Case 4: Licensing or Feature Toggles
            //   - Store a hidden link to a license file or configuration workbook.
            //     The application reads the external file at runtime while the
            //     link remains invisible in the Excel UI.

            // Use Case 5: Security Through Obscurity
            //   - Hide sensitive external references (e.g., financial models) from
            //     casual users who might otherwise see the file paths in visible
            //     hyperlinks.

            // ------------------------------------------------------------
            // 4. Clean up (optional)
            // ------------------------------------------------------------
            wb.Dispose();
            loadedWb.Dispose();
        }
    }
}