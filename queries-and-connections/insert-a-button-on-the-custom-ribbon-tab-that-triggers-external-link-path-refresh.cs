using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsRibbonButtonDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑enabled format is required for the button to call a macro)
            Workbook workbook = new Workbook();

            // ------------------------------------------------------------
            // 1. Add an external link that we want to refresh later.
            // ------------------------------------------------------------
            // Example: link to an external workbook named "DataSource.xlsx" with two sheets.
            string externalFile = "DataSource.xlsx";
            string[] sheetNames = new string[] { "Sheet1", "Sheet2" };
            // Add the external link to the workbook's ExternalLinks collection.
            int linkIndex = workbook.Worksheets.ExternalLinks.Add(externalFile, sheetNames);
            // (Optional) Store the index for later reference or manipulation.
            ExternalLink externalLink = workbook.Worksheets.ExternalLinks[linkIndex];

            // ------------------------------------------------------------
            // 2. Define custom Ribbon XML with a button.
            // ------------------------------------------------------------
            // The button calls a macro named "RefreshExternalLinks". The macro must be
            // present in the workbook (added via VBA) to perform the actual refresh.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tools\">" +
                "        <group id=\"linkGroup\" label=\"External Links\">" +
                "          <button id=\"refreshButton\" " +
                "                  label=\"Refresh Links\" " +
                "                  size=\"large\" " +
                "                  onAction=\"RefreshExternalLinks\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook.
            workbook.RibbonXml = ribbonXml;

            // ------------------------------------------------------------
            // 3. (Optional) Add a simple shape button on the worksheet for UI testing.
            // ------------------------------------------------------------
            // This demonstrates the ShapeCollection.AddButton method.
            Worksheet sheet = workbook.Worksheets[0];
            Button sheetButton = sheet.Shapes.AddButton(topRow: 2, top: 0, leftColumn: 2, left: 0, height: 30, width: 120);
            sheetButton.Text = "Refresh Links (Sheet Button)";
            // Link the shape button to the same macro.
            sheetButton.MacroName = "RefreshExternalLinks";

            // ------------------------------------------------------------
            // 4. Save the workbook as a macro‑enabled file.
            // ------------------------------------------------------------
            workbook.Save("WorkbookWithCustomRibbon.xlsm");

            // Inform the user.
            Console.WriteLine("Workbook created with custom ribbon button and external link.");
        }
    }
}