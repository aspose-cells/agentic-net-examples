using System;
using Aspose.Cells;

namespace RibbonWorkbookVerification
{
    class Program
    {
        static void Main()
        {
            // Define the output file path
            string outputPath = "CustomRibbonWorkbook.xlsm";

            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Sample ribbon XML to customize the ribbon UI
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

            // Apply the custom ribbon XML to the workbook
            workbook.RibbonXml = ribbonXml;

            // Save the workbook to disk (lifecycle: save)
            workbook.Save(outputPath, SaveFormat.Xlsm);

            // Attempt to load the saved workbook (lifecycle: load) and verify no errors
            try
            {
                Workbook loadedWorkbook = new Workbook(outputPath);
                Console.WriteLine("Workbook loaded successfully: " + (loadedWorkbook != null));

                // Optional: verify that the RibbonXml property is retained
                bool ribbonSet = !string.IsNullOrEmpty(loadedWorkbook.RibbonXml);
                Console.WriteLine("RibbonXml retained after load: " + ribbonSet);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading workbook: " + ex.Message);
            }
        }
    }
}