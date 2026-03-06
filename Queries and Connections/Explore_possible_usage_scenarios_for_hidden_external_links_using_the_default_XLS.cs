using System;
using Aspose.Cells;

namespace AsposeCellsExternalLinkDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add an external link
            // -------------------------------------------------
            Workbook wb = new Workbook();

            // Prepare external file name and the sheet names to reference
            string externalFile = "ExternalData.xlsx";
            string[] sheetNames = new string[] { "Sheet1" };

            // Add the external link to the workbook's external links collection
            // This will also create a hidden external link entry (IsVisible = false)
            int linkIndex = wb.Worksheets.ExternalLinks.Add(externalFile, sheetNames);

            // Optionally, add a formula that uses the external link
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].Formula = $"='[{externalFile}]Sheet1'!A1";

            // Save the workbook in the default XLSX format
            wb.Save("HiddenExternalLinkDemo.xlsx");

            // -------------------------------------------------
            // 2. Load the saved workbook and inspect external links
            // -------------------------------------------------
            Workbook loadedWb = new Workbook("HiddenExternalLinkDemo.xlsx");

            // Iterate through all external links and display their properties
            Console.WriteLine("External links in the loaded workbook:");
            foreach (ExternalLink link in loadedWb.Worksheets.ExternalLinks)
            {
                // DataSource – the path to the external file
                Console.WriteLine($"  DataSource : {link.DataSource}");

                // IsVisible – indicates whether the link is visible in Excel UI
                // For links added programmatically without user interaction, this is typically false
                Console.WriteLine($"  IsVisible  : {link.IsVisible}");

                // Additional useful properties
                Console.WriteLine($"  IsReferred : {link.IsReferred}");
                Console.WriteLine($"  PathType   : {link.PathType}");
                Console.WriteLine();
            }

            // -------------------------------------------------
            // 3. Demonstrate a scenario where a hidden external link becomes visible
            // -------------------------------------------------
            // In Excel, a hidden external link can become visible when a user explicitly
            // creates a defined name that references the external source.
            // We simulate this by adding an external name to the link.

            if (loadedWb.Worksheets.ExternalLinks.Count > 0)
            {
                ExternalLink firstLink = loadedWb.Worksheets.ExternalLinks[0];

                // Add an external name – this operation does not change IsVisible directly,
                // but it makes the link discoverable when the workbook is opened in Excel.
                firstLink.AddExternalName("MyExternalValue", "Sheet1!A1");

                // Save the workbook again to persist the change
                loadedWb.Save("HiddenExternalLinkDemo_Updated.xlsx");

                // Reload and show that the link still reports IsVisible = false
                // (Excel treats the link as visible only when the user accesses it via UI)
                Workbook reloaded = new Workbook("HiddenExternalLinkDemo_Updated.xlsx");
                ExternalLink reloadedLink = reloaded.Worksheets.ExternalLinks[0];
                Console.WriteLine("After adding an external name:");
                Console.WriteLine($"  DataSource : {reloadedLink.DataSource}");
                Console.WriteLine($"  IsVisible  : {reloadedLink.IsVisible}");
            }

            // -------------------------------------------------
            // 4. Clean up
            // -------------------------------------------------
            wb.Dispose();
            loadedWb.Dispose();
        }
    }
}