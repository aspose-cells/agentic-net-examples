using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CustomRibbonDataToolsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Define custom Ribbon XML for the "Data Tools" tab (idMso="TabData")
                string ribbonXml =
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                    "  <ribbon>" +
                    "    <tabs>" +
                    "      <tab idMso=\"TabData\">" +
                    "        <group id=\"ConnectionManagementGroup\" label=\"Connection Management\">" +
                    "          <button id=\"RefreshConnection\" label=\"Refresh\" size=\"large\" imageMso=\"RefreshAll\" />" +
                    "          <button id=\"EditConnection\" label=\"Edit Connection\" size=\"large\" imageMso=\"DataConnectionProperties\" />" +
                    "        </group>" +
                    "      </tab>" +
                    "    </tabs>" +
                    "  </ribbon>" +
                    "</customUI>";

                // Assign the custom UI to the workbook
                workbook.RibbonXml = ribbonXml;

                // Save as macro‑enabled workbook (required for custom UI)
                workbook.Save("CustomRibbonDataTools.xlsm");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomRibbonDataToolsDemo.Run();
        }
    }
}