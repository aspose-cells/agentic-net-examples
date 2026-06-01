using Aspose.Cells;
using System;

class RibbonShortcutDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML with a custom button.
        // The keyTip attribute defines the keyboard shortcut (Alt + C in this case).
        // The onAction attribute can point to a macro name (optional for this demo).
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"My Tab\">" +
            "        <group id=\"customGroup\" label=\"My Group\">" +
            "          <button id=\"customButton\" label=\"My Button\" size=\"large\" keyTip=\"C\" onAction=\"MyMacro\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Save the workbook (macro-enabled format to retain the custom UI)
        workbook.Save("CustomRibbonWithShortcut.xlsm");
    }
}