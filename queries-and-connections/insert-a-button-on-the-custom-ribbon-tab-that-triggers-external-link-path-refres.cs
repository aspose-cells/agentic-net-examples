using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

namespace AsposeCellsRibbonButtonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑enabled format)
            Workbook workbook = new Workbook();

            // -----------------------------------------------------------------
            // 1. Define custom Ribbon XML with a button that calls a macro.
            //    The button is placed on a custom tab and group.
            // -----------------------------------------------------------------
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"Link Tools\">" +
                "        <group id=\"linkGroup\" label=\"External Links\">" +
                "          <button id=\"refreshLinksBtn\" " +
                "                  label=\"Refresh Links\" " +
                "                  size=\"large\" " +
                "                  onAction=\"RefreshLinks\" />" + // Calls VBA macro named RefreshLinks
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // -----------------------------------------------------------------
            // 2. Add an example external link (to demonstrate refresh settings)
            // -----------------------------------------------------------------
            string externalFile = @"C:\Data\ExternalData.xlsx";
            string[] sheets = new string[] { "Sheet1", "Sheet2" };
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(externalFile, sheets);

            // -----------------------------------------------------------------
            // 3. Configure the external connection to refresh when the file opens.
            //    This ensures the link is refreshed automatically (the macro can
            //    also invoke a manual refresh if needed).
            // -----------------------------------------------------------------
            if (workbook.DataConnections.Count > 0)
            {
                ExternalConnection conn = workbook.DataConnections[0];
                conn.RefreshOnLoad = true;          // Refresh on workbook open
                conn.BackgroundRefresh = false;    // Synchronous refresh
            }

            // -----------------------------------------------------------------
            // 4. Save the workbook as a macro‑enabled file.
            // -----------------------------------------------------------------
            workbook.Save("RibbonRefreshLinksDemo.xlsm");
        }
    }
}