// Title: Add a Data Tools ribbon tab to an Excel workbook with Aspose.Cells for .NET
// Description: Shows how to create a macro‑enabled .xlsm file, inject custom Ribbon XML that defines a “Data Tools” tab with a large Refresh button, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | RibbonXml | custom ribbon tab | Data Tools tab | Excel .xlsm | C# | macro‑enabled workbook | custom UI | Excel add‑in generation | programmatic ribbon
// Common Searches: Aspose.Cells add custom ribbon tab C# | How to set RibbonXml in Aspose.Cells | Create Excel file with custom UI using Aspose.Cells | Save workbook with custom ribbon as .xlsm | Define ribbon XML for Excel in .NET
// Developer Intent: Programmatically embed a custom ribbon tab in an Excel workbook and preserve it in a macro‑enabled file.
// Use Cases: Provide end‑users with a dedicated Data Tools tab that launches data‑refresh macros. | Distribute workbooks that already contain a predefined UI for import/export operations. | Automate generation of Excel add‑ins with custom ribbon groups for corporate reporting tools.
// AI Prompts: Generate C# code to add additional buttons to the Data Tools ribbon tab with Aspose.Cells. | Show how to link the Refresh button in the custom ribbon to a VBA macro using Aspose.Cells. | Explain how to extend the RibbonXml to include a dropdown list and toggle controls in the Data Tools group.

using System;
using Aspose.Cells;

// Shows how to create a macro‑enabled .xlsm file, inject custom Ribbon XML that defines a “Data Tools” tab with a large Refresh button, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Define the Ribbon XML that adds a custom tab named "Data Tools"
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"dataToolsTab\" label=\"Data Tools\">" +
            "        <group id=\"dataToolsGroup\" label=\"Data Tools Group\">" +
            "          <button id=\"btnRefresh\" label=\"Refresh\" size=\"large\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (use a macro-enabled format to preserve the custom UI)
        workbook.Save("DataToolsRibbon.xlsm");
    }
}
