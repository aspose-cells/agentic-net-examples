using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RibbonXmlUpdateDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define Ribbon XML using the newer namespace (Office 2010+)
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
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

            // Assign the Ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Set OOXML compliance to strict mode for better future compatibility
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Define output file path
            string outputPath = "UpdatedRibbonWorkbook.xlsm";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a macro-enabled file (required for custom UI)
            workbook.Save(outputPath);

            // Confirmation output
            Console.WriteLine("Ribbon XML updated and workbook saved.");
        }
    }
}