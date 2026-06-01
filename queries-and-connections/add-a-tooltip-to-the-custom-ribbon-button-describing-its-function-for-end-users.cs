using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    public class AddRibbonTooltipDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define ribbon XML with a tooltip (ScreenTip) for the custom button
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" screentip=\"Executes the custom operation\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";

            // Assign the custom ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Define output file path
            string outputPath = "CustomRibbonWithTooltip.xlsm";

            // Save the workbook (macro-enabled to retain ribbon UI)
            workbook.Save(outputPath);
        }
    }
}