using System;
using Aspose.Cells;

namespace AsposeCellsRibbonShortcutDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Define Ribbon XML with a custom button that includes a keytip.
            // The keytip attribute defines the keyboard shortcut that appears after pressing Alt.
            // Here, pressing Alt then 'A' will activate the button.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" " +
                "                  label=\"My Button\" " +
                "                  size=\"large\" " +
                "                  keytip=\"A\" />" + // Keyboard shortcut Alt+A
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the Ribbon XML to the workbook (feature: RibbonXml property)
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (lifecycle: save)
            workbook.Save("CustomRibbonWithShortcut.xlsm");

            Console.WriteLine("Workbook with custom ribbon button and shortcut created successfully.");
        }
    }
}