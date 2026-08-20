// Title: Add a Custom Ribbon Tab with a Refresh Links Button & Auto‑Refresh External Connections (Aspose.Cells .NET)
// Description: Creates a new Workbook, injects Ribbon XML that adds a custom tab with a large "Refresh Links" button linked to a VBA macro, sets every ExternalConnection.RefreshOnLoad to true, and saves the file as a macro‑enabled .xlsm workbook.
// Keywords: Aspose.Cells custom ribbon | C# add ribbon button | RefreshLinks VBA macro | external connections RefreshOnLoad | macro enabled .xlsm with custom UI | Excel custom UI XML | auto refresh external data Aspose.Cells | Aspose.Cells RibbonXml property
// Common Searches: Aspose.Cells add custom ribbon tab C# | How to create a Refresh Links button in Excel with Aspose.Cells | Enable automatic refresh of external connections in Aspose.Cells workbook | Save workbook as macro enabled .xlsm using Aspose.Cells | Inject custom UI XML into Excel file with Aspose.Cells
// Developer Intent: Generate a macro‑enabled workbook that includes a custom ribbon button to invoke a RefreshLinks macro and automatically refreshes all external data connections on load.
// Use Cases: Provide end users a one‑click button on a custom ribbon tab to update linked data sources. | Ensure external connections are always current by enabling RefreshOnLoad for every connection. | Distribute a ready‑to‑use .xlsm file that combines custom UI and auto‑refresh for non‑technical stakeholders.
// AI Prompts: Write the VBA macro RefreshLinks that loops through all workbook connections and calls their Refresh method. | Show how to add additional buttons to the same custom ribbon group using the RibbonXml property in Aspose.Cells. | Explain how to merge new RibbonXml with existing custom UI without overwriting previous definitions.

using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

// Creates a new Workbook, injects Ribbon XML that adds a custom tab with a large "Refresh Links" button linked to a VBA macro, sets every ExternalConnection.RefreshOnLoad to true, and saves the file as a macro‑enabled .xlsm workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define custom Ribbon XML that adds a button to a custom tab.
        // The button's onAction attribute points to a macro named "RefreshLinks".
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"My Tools\">" +
            "        <group id=\"linkGroup\" label=\"Links\">" +
            "          <button id=\"refreshBtn\" label=\"Refresh Links\" size=\"large\" onAction=\"RefreshLinks\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Configure all external connections to refresh when the workbook is opened.
        foreach (ExternalConnection connection in workbook.DataConnections)
        {
            connection.RefreshOnLoad = true;
        }

        // Save the workbook as a macro‑enabled file (the button expects a VBA macro named RefreshLinks).
        workbook.Save("WorkbookWithCustomRibbon.xlsm");
    }
}
