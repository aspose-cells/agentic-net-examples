// Title: Verify that a workbook saved with custom Ribbon XML opens without errors using Aspose.Cells for .NET
// AI Prompts: Write a C# console app that creates a new Workbook, assigns a custom Ribbon XML definition, saves it as an .xlsm file, reloads the file, and asserts that the RibbonXml property is still populated. | Generate code that demonstrates how to persist a custom UI ribbon in an Aspose.Cells workbook, then load the workbook to confirm no exceptions and that the ribbon definition is retained.
// Common Searches: how to persist custom ribbon xml in an Aspose.Cells .xlsm workbook and reload it | Aspose.Cells C# verify RibbonXml property after saving and opening workbook | example of saving and loading a workbook with custom UI ribbon using Aspose.Cells for .NET | check if custom ribbon definition is retained in saved Excel macro-enabled file with Aspose.Cells | C# load workbook with RibbonXml without throwing errors Aspose.Cells
// Tags: custom ribbon xml Aspose.Cells | save workbook with ribbon xml .xlsm | load workbook and verify RibbonXml | C# Aspose.Cells ribbon persistence validation | macro-enabled workbook custom UI verification

using System;
using Aspose.Cells;

// Creates a new workbook, applies custom Ribbon XML, saves it as an .xlsm file, reloads the file, and confirms that the RibbonXml property remains intact.
class VerifyWorkbookRibbon
{
    static void Main()
    {
        // Create a new workbook instance
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

        // Apply the Ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Specify the file path for saving
        string filePath = "CustomRibbonWorkbook.xlsm";

        // Save the workbook to disk
        workbook.Save(filePath);

        // Load the saved workbook to verify it opens without errors
        Workbook loadedWorkbook = new Workbook(filePath);

        // Check that the RibbonXml property is retained after loading
        bool ribbonIsSet = !string.IsNullOrEmpty(loadedWorkbook.RibbonXml);
        Console.WriteLine("RibbonXml loaded successfully: " + ribbonIsSet);
    }
}
