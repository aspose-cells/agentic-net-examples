using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RibbonButtonTooltipDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML with a tooltip (supertip) for the button
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

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook (as a macro-enabled file to retain the ribbon UI)
            workbook.Save("WorkbookWithRibbonTooltip.xlsm");

            // Optional: verify that RibbonXml is set
            Console.WriteLine("RibbonXml set: " + (workbook.RibbonXml != null));
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RibbonButtonTooltipDemo.Run();
        }
    }
}