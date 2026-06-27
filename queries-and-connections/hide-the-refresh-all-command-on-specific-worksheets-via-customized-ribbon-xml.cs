using System;
using Aspose.Cells;

namespace AsposeCellsRibbonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add additional worksheets (optional)
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Ribbon XML that disables the built‑in "Refresh All" command (idMso="RefreshAll")
            // This XML is applied to the whole workbook; per‑sheet control would require callbacks,
            // which are beyond the scope of this simple example.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <commands>" +
                "    <command idMso=\"RefreshAll\" enabled=\"false\" />" +
                "  </commands>" +
                "</customUI>";

            // Set the custom Ribbon XML
            workbook.RibbonXml = ribbonXml;

            // Save as a macro‑enabled workbook so the Ribbon customization is preserved
            workbook.Save("HideRefreshAllDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}