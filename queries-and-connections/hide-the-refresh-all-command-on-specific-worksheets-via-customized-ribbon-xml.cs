// Title: Hide the Refresh All button on selected worksheets using custom Ribbon XML with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, adds a worksheet, assigns RibbonXml to disable and hide the built‑in RefreshAll command, and saves the file as an XLSM. | Explain how to embed custom Ribbon XML via the Workbook.RibbonXml property to remove the Refresh All control from the Excel UI, and why the workbook must be saved as a macro‑enabled format. | Show how to modify the <commands> element in Ribbon XML to hide a specific built‑in Excel command for particular sheets using Aspose.Cells for .NET.
// Common Searches: how to hide refresh all button in Excel using Aspose.Cells C# | custom ribbon xml hide built‑in commands Aspose.Cells .NET example | disable specific Excel ribbon controls for certain worksheets programmatically | Aspose.Cells set RibbonXml property to remove RefreshAll command | save workbook with custom UI as macro enabled file Aspose.Cells
// Tags: custom ribbon XML for Excel workbook Aspose.Cells | disable built‑in ribbon controls programmatically | macro‑enabled XLSM workbook with custom UI | assign RibbonXml property in .NET | modify commands element in Ribbon XML

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds a second worksheet, defines custom Ribbon XML that disables and hides the built‑in RefreshAll command, assigns this XML to the workbook's RibbonXml property, and saves the workbook as a macro‑enabled XLSM file.
class HideRefreshAllRibbonDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet (example of a sheet where you might want the command hidden)
        workbook.Worksheets.Add("Sheet2");

        // Ribbon XML that hides the built‑in "Refresh All" command.
        // The <commands> element allows us to modify built‑in controls.
        // Setting both enabled and visible to false removes the command from the UI.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <commands>" +
            "    <command idMso=\"RefreshAll\" enabled=\"false\" visible=\"false\" />" +
            "  </commands>" +
            "</customUI>";

        // Apply the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro‑enabled file (XLSM) because custom UI requires it
        workbook.Save("HideRefreshAllDemo.xlsm", SaveFormat.Xlsm);
    }
}
