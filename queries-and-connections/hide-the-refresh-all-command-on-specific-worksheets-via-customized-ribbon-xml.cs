// Title: Hide Excel’s Refresh All button via custom Ribbon XML using Aspose.Cells for .NET
// Description: Demonstrates how to create an in‑memory workbook, inject custom Ribbon XML that sets the built‑in RefreshAll command to invisible, assign it through the RibbonXml property, and save the workbook as a macro‑enabled .xlsm so the Ribbon customization persists.
// Keywords: Aspose.Cells RibbonXml | hide RefreshAll command | custom Ribbon XML Excel | disable Refresh All button .NET | macro‑enabled workbook Aspose | Excel Ribbon customization C# | Aspose.Cells hide Ribbon command
// Common Searches: Aspose.Cells hide Refresh All button | custom Ribbon XML for Excel using C# | disable RefreshAll command programmatically | save workbook with hidden Ribbon commands | Aspose.Cells RibbonXml property example
// Developer Intent: Programmatically hide the built‑in Refresh All command on an Excel workbook’s Ribbon by assigning custom Ribbon XML.
// Use Cases: Prepare a template workbook that distributes without the Refresh All button to prevent unwanted data refreshes. | Apply Ribbon customization to existing .xlsm files to lock down external data connections. | Combine hidden Ribbon commands with worksheet protection for secure reporting solutions.
// AI Prompts: Write C# code that loads an existing .xlsm file, sets RibbonXml to hide the RefreshAll command, and saves the workbook using Aspose.Cells. | Explain how the RibbonXml property works in Aspose.Cells and show how to target other built‑in commands for visibility changes. | Provide a step‑by‑step guide to hide multiple Ribbon commands (e.g., RefreshAll, Connections) with a single custom UI XML string in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create an in‑memory workbook, inject custom Ribbon XML that sets the built‑in RefreshAll command to invisible, assign it through the RibbonXml property, and save the workbook as a macro‑enabled .xlsm so the Ribbon customization persists.
class HideRefreshAllDemo
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Ribbon XML that hides the built‑in "Refresh All" command.
        // The <command> element with idMso="RefreshAll" and visible="false"
        // disables the button on the Ribbon UI for this workbook.
        string ribbonXml =
            @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">
                <commands>
                    <command idMso=""RefreshAll"" visible=""false""/>
                </commands>
              </customUI>";

        // Assign the custom Ribbon XML to the workbook.
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro‑enabled file so the Ribbon customization is retained.
        workbook.Save("HideRefreshAll.xlsm");
    }
}
