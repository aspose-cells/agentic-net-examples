using System;
using Aspose.Cells;

namespace RibbonWorkbookVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Define custom Ribbon XML
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

            // Apply the custom Ribbon XML to the workbook (RibbonXml property)
            workbook.RibbonXml = ribbonXml;

            // Save the workbook to a temporary file (uses the Save(string) method rule)
            string tempPath = "TempRibbonWorkbook.xlsm";
            workbook.Save(tempPath, SaveFormat.Xlsm);

            // Attempt to load the saved workbook to verify it opens without errors
            try
            {
                // Load the workbook from the saved path (uses the Workbook(string) constructor rule)
                Workbook loadedWorkbook = new Workbook(tempPath);

                // If we reach this point, the workbook opened successfully
                Console.WriteLine("Workbook loaded successfully. RibbonXml is set: " +
                                  (loadedWorkbook.RibbonXml != null));

                // Optional: clean up
                loadedWorkbook.Dispose();
            }
            catch (Exception ex)
            {
                // If an exception occurs, report the failure
                Console.WriteLine("Failed to load workbook: " + ex.Message);
            }
            finally
            {
                // Dispose the original workbook
                workbook.Dispose();
            }
        }
    }
}