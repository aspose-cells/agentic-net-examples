// Title: How to add a supertip tooltip to a custom ribbon button in an Excel .xlsm workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, defines RibbonXml with a button that includes a supertip attribute, and saves the file as a macro‑enabled .xlsm. | Show how to modify existing RibbonXml to insert a supertip tooltip for a custom ribbon button with Aspose.Cells. | Provide a step‑by‑step example of assigning RibbonXml containing a button with a supertip to a workbook and persisting the ribbon customization.
// Common Searches: Aspose.Cells C# add supertip to custom ribbon button in .xlsm file | how to set tooltip for Excel ribbon button using RibbonXml property | example of custom ribbon XML with tooltip for macro‑enabled workbook in Aspose.Cells | C# code to create custom ribbon tab with button tooltip in Excel using Aspose.Cells
// Tags: custom ribbon button supertip Aspose.Cells | RibbonXml tooltip attribute C# | save macro-enabled workbook with ribbon customizations | Aspose.Cells add tooltip to ribbon UI | Excel .xlsm custom ribbon XML example

using System;
using Aspose.Cells;

// The example creates a new Workbook, builds custom ribbon XML that defines a button with a supertip tooltip, assigns the XML to the workbook's RibbonXml property, and saves the workbook as a macro‑enabled .xlsm file to retain the ribbon customization.
class RibbonTooltipDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define custom ribbon XML with a tooltip (supertip) for the button
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"My Tab\">" +
            "        <group id=\"customGroup\" label=\"My Group\">" +
            "          <button id=\"customButton\" label=\"My Button\" size=\"large\" supertip=\"Executes my custom action\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the XML to the workbook's RibbonXml property
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (macro-enabled format to retain ribbon customizations)
        workbook.Save("CustomRibbonWithTooltip.xlsm");
    }
}
