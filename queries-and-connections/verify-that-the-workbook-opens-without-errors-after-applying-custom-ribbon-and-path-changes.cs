using System;
using Aspose.Cells;

class VerifyWorkbookRibbon
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define custom ribbon XML
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

        // Apply the ribbon XML to the workbook
        workbook.RibbonXml = ribbonXml;

        // Change the absolute path where the workbook will be saved
        workbook.AbsolutePath = @"C:\Temp\CustomRibbonWorkbook.xlsm";

        // Save the workbook (using the lifecycle save rule)
        workbook.Save(workbook.AbsolutePath, SaveFormat.Xlsm);

        // Load the saved workbook to verify it opens without errors (using the lifecycle load rule)
        Workbook loadedWorkbook = new Workbook(workbook.AbsolutePath);

        // Verify that the RibbonXml property is retained after loading
        bool ribbonRetained = loadedWorkbook.RibbonXml != null && loadedWorkbook.RibbonXml.Contains("customTab");
        Console.WriteLine("RibbonXml retained after load: " + ribbonRetained);

        // Dispose resources
        workbook.Dispose();
        loadedWorkbook.Dispose();
    }
}