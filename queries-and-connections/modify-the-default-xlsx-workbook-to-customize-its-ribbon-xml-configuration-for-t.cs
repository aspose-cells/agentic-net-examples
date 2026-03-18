using System;
using Aspose.Cells;

namespace RibbonCustomizationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML to add a new tab, group, and button
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Set the RibbonXml property of the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook as a macro-enabled file to retain the Ribbon UI
            workbook.Save("CustomizedRibbon.xlsm");

            // Verify that the RibbonXml property has been set
            Console.WriteLine("RibbonXml is set: " + (workbook.RibbonXml != null));
        }
    }
}