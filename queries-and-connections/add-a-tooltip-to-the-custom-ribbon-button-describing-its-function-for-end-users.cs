// Title: Add a SuperTip tooltip to a custom ribbon button in an Excel workbook with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, defines Ribbon XML that includes a large button with a supertip attribute, assigns the XML to the workbook's RibbonXml property, and saves the file as a macro‑enabled .xlsm to preserve the custom ribbon UI.
// Keywords: Aspose.Cells | RibbonXml | tooltip | supertip | custom ribbon | C# | Excel workbook | macro-enabled | UI customization
// Common Searches: Aspose.Cells add tooltip to custom ribbon button | RibbonXml supertip example C# | How to set a tooltip for an Excel custom UI button using Aspose.Cells | Create macro-enabled workbook with custom ribbon and tooltip | C# code for custom ribbon tooltip in Excel
// Developer Intent: Add a supertip tooltip to a custom ribbon button in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Guide end users by displaying a descriptive tooltip on a custom ribbon button. | Improve accessibility of custom Excel UI elements with concise supertips. | Reuse a RibbonXml template containing tooltips across multiple workbooks for consistent UI hints.
// AI Prompts: Generate C# code that adds a supertip tooltip to a custom ribbon button in an Aspose.Cells workbook and saves it as .xlsm. | Show how to modify existing RibbonXml to include tooltips for several custom ribbon buttons with Aspose.Cells for .NET. | Explain how to localize the supertip text of a custom ribbon button in an Aspose.Cells workbook. | Provide a step‑by‑step guide to embed a custom ribbon with tooltips into a macro‑enabled Excel file using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new Workbook, defines Ribbon XML that includes a large button with a supertip attribute, assigns the XML to the workbook's RibbonXml property, and saves the file as a macro‑enabled .xlsm to preserve the custom ribbon UI.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML with a button that includes a tooltip (supertip)
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

        // Assign the custom Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (macro-enabled format to retain Ribbon UI)
        workbook.Save("CustomRibbonWithTooltip.xlsm");
    }
}
