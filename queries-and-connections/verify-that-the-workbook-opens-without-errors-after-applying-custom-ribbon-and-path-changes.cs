// Title: Aspose.Cells .NET – Verify Custom Ribbon XML Persists When Reloading an XLSM Workbook
// Description: Creates a workbook, assigns custom Ribbon XML, saves it as an XLSM file, and reloads the file to confirm the Ribbon UI is retained and no exceptions are thrown.
// Keywords: Aspose.Cells custom ribbon | RibbonXml property | save XLSM with UI | reload workbook after ribbon changes | verify workbook opens | C# Aspose.Cells example
// Common Searches: add custom ribbon xml Aspose.Cells .NET | save and load XLSM with custom UI using Aspose.Cells | check RibbonXml after workbook reload | Aspose.Cells verify workbook opens without error | C# example custom ribbon Excel file
// Developer Intent: Ensure that a workbook saved with a custom Ribbon XML definition can be opened again without errors and that the RibbonXml property remains populated.
// Use Cases: Inject a custom Ribbon UI into a newly created workbook for branding or automation. | Persist the custom UI in an XLSM file and validate its presence after saving. | Detect and handle any loading issues caused by embedded Ribbon XML.
// AI Prompts: Write C# code that adds custom Ribbon XML to an Excel workbook with Aspose.Cells, saves it as .xlsm, and confirms the file loads without exceptions. | Explain how to test the persistence of the RibbonXml property after saving and reopening a workbook using Aspose.Cells for .NET. | Provide a step‑by‑step guide for handling errors when loading an XLSM file that contains custom Ribbon UI.

using System;
using Aspose.Cells;

namespace RibbonWorkbookTest
{
    // Creates a workbook, assigns custom Ribbon XML, saves it as an XLSM file, and reloads the file to confirm the Ribbon UI is retained and no exceptions are thrown.
    class Program
    {
        static void Main()
        {
            // Define the output file path
            string filePath = "CustomRibbonWorkbook.xlsm";

            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Custom ribbon XML to be applied
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

            // Save the workbook with the custom ribbon
            workbook.Save(filePath, SaveFormat.Xlsm);

            // Attempt to load the saved workbook to verify it opens without errors
            try
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully. RibbonXml is set: " + (loadedWorkbook.RibbonXml != null));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error loading workbook: " + ex.Message);
            }
        }
    }
}
