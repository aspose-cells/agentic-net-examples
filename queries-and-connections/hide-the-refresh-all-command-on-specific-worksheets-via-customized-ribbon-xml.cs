using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");

        // Ribbon XML that hides the built‑in "Refresh All" command (idMso="RefreshAll")
        // The command is set to visible="false" so it will not appear in the Ribbon UI.
        string ribbonXml =
            @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" +
            @"  <ribbon>" +
            @"    <commands>" +
            @"      <command idMso=""RefreshAll"" visible=""false""/>" +
            @"    </commands>" +
            @"  </ribbon>" +
            @"</customUI>";

        // Apply the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (macro‑enabled format is not required for this customization)
        workbook.Save("HideRefreshAll.xlsx", SaveFormat.Xlsx);
    }
}