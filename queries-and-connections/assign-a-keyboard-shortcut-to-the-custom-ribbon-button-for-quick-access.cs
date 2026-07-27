using Aspose.Cells;

class RibbonShortcutDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define Ribbon XML with a button that has a keyTip (keyboard shortcut after pressing Alt)
        string ribbonXml =
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
            "  <ribbon>" +
            "    <tabs>" +
            "      <tab id=\"customTab\" label=\"My Tab\">" +
            "        <group id=\"customGroup\" label=\"My Group\">" +
            "          <button id=\"customButton\" label=\"My Button\" size=\"large\" keyTip=\"C\" />" +
            "        </group>" +
            "      </tab>" +
            "    </tabs>" +
            "  </ribbon>" +
            "</customUI>";

        // Assign the Ribbon XML to the workbook (uses Workbook.RibbonXml property)
        workbook.RibbonXml = ribbonXml;

        // Save the workbook as a macro-enabled file to preserve the custom UI
        workbook.Save("CustomRibbonWithShortcut.xlsm");
    }
}