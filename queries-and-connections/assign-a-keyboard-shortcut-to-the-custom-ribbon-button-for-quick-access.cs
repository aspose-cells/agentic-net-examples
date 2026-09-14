// Title: Assign a keytip keyboard shortcut to a custom ribbon button in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that creates a macro-enabled workbook, defines custom ribbon XML with a button that includes a keytip, and saves the file. | Demonstrate how to set the RibbonXml property in Aspose.Cells to embed a keytip for quick keyboard access to a custom ribbon UI element.
// Common Searches: asp.net cells add keytip to custom ribbon button c# example | c# assign Alt+F shortcut to custom Excel ribbon button using Aspose.Cells | how to embed custom ribbon XML with keyboard shortcut in a macro-enabled workbook Aspose.Cells | set keytip attribute in RibbonXml for Aspose.Cells workbook c#
// Tags: custom ribbon keytip Aspose.Cells C# | RibbonXml keyboard shortcut macro-enabled workbook | Aspose.Cells add custom UI button with keytip | C# set keytip attribute in RibbonXml | Excel ribbon shortcut via Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsRibbonShortcutDemo
{
    // The example creates a new Workbook, builds Ribbon XML that defines a custom tab, group, and a large button with the keytip "C" (triggered by Alt+F, then the tab, then C), assigns this XML to the workbook's RibbonXml property, and saves the result as a macro-enabled .xlsm file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define Ribbon XML with a custom button that includes a keytip (keyboard shortcut)
            // The keytip "C" means the user can press Alt+F, then the custom tab, then "C" to activate the button.
            // Adjust the keytip value as needed for your shortcut.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" keytip=\"C\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file (required for custom ribbon UI)
            workbook.Save("CustomRibbonWithShortcut.xlsm");
        }
    }
}
