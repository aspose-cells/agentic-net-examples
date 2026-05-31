using System;
using Aspose.Cells;

class HideRefreshAllRibbonDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add worksheets (example: Sheet1, Sheet2, Sheet3)
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Ribbon XML that hides the built‑in "Refresh All" command
        // The XML is applied at the workbook level; the command will be hidden
        // for all sheets (custom logic per sheet would require VBA callbacks)
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"TabData\">" +
            "        <group idMso=\"GroupRefresh\">" +
            "          <control idMso=\"RefreshAll\" visible=\"false\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro‑enabled file (required for custom UI)
        workbook.Save("HideRefreshAllDemo.xlsm", SaveFormat.Xlsm);
    }
}