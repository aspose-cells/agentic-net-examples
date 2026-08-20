// Title: Add a custom "Connection Management" ribbon group to the Data tab using Aspose.Cells for .NET
// Description: C# example that uses Aspose.Cells to inject RibbonXml into a workbook, creating a new "Connection Management" group on the built‑in Data tab (idMso="TabData"). The group contains large Refresh and Edit Connection buttons with onAction callbacks, and the workbook is saved as a macro‑enabled .xlsm to preserve the custom UI.
// Keywords: Aspose.Cells RibbonXml | custom ribbon group | Data tab Excel | connection management buttons | C# Aspose.Cells example | macro enabled workbook | Excel custom UI | add ribbon button programmatically
// Common Searches: Aspose.Cells add custom ribbon group to Data tab | how to create RibbonXml with connection buttons in C# | save custom ribbon in .xlsm using Aspose.Cells | Excel custom UI for connection refresh button | programmatically extend Excel Data Tools ribbon
// Developer Intent: Generate a workbook that embeds a custom ribbon group on the Data tab, providing Refresh and Edit Connection commands via Aspose.Cells.
// Use Cases: Distribute workbooks that let end‑users manage data connections directly from the ribbon. | Enhance existing Excel templates with connection‑related shortcuts without manual UI editing. | Create macro‑enabled files that retain custom UI for corporate reporting tools.
// AI Prompts: Write C# code with Aspose.Cells to add a third "Delete Connection" button to the custom ribbon group and implement its onAction handler. | Explain how to load an existing .xlsm file, modify its RibbonXml to include a new group, and save the changes with Aspose.Cells. | Provide a step‑by‑step guide to test the custom ribbon in Excel, covering macro security settings and button callback verification.

using System;
using Aspose.Cells;

// C# example that uses Aspose.Cells to inject RibbonXml into a workbook, creating a new "Connection Management" group on the built‑in Data tab (idMso="TabData"). The group contains large Refresh and Edit Connection buttons with onAction callbacks, and the workbook is saved as a macro‑enabled .xlsm to preserve the custom UI.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty workbook)
        Workbook workbook = new Workbook();

        // Define custom Ribbon XML.
        // The XML adds a new group called "Connection Management" inside the built‑in "Data" tab (idMso="TabData").
        // Two sample buttons are added: Refresh and Edit Connection.
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab idMso=\"TabData\">" +                     // Existing Data Tools tab
            "        <group id=\"connectionGroup\" label=\"Connection Management\">" +
            "          <button id=\"refreshBtn\" label=\"Refresh\" size=\"large\" onAction=\"RefreshConnections\" />" +
            "          <button id=\"editBtn\"   label=\"Edit Connection\" size=\"large\" onAction=\"EditConnection\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the custom UI to the workbook.
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro‑enabled file (required for custom UI to be retained).
        workbook.Save("ConnectionRibbon.xlsm");
    }
}
