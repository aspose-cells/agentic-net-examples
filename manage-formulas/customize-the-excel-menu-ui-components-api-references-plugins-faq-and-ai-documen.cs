using System;
using Aspose.Cells;

namespace ExcelCustomizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path where the customized workbook will be saved
            string outputPath = "customized_output.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // 1. Customize the Ribbon UI by setting the RibbonXml property.
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\">" +
                "  <ribbon>" +
                "    <tabs>" +
                "      <tab id=\"customTab\" label=\"My Custom Tab\">" +
                "        <group id=\"customGroup\" label=\"My Group\">" +
                "          <button id=\"customButton\" label=\"My Button\" size=\"large\" />" +
                "        </group>" +
                "      </tab>" +
                "    </tabs>" +
                "  </ribbon>" +
                "</customUI>";
            workbook.RibbonXml = ribbonXml;

            // 2. Register an add‑in function (plugin) into the workbook.
            string addInFilePath = "MyAddIn.xlam";
            string functionName = "MY_CUSTOM_UDF";

            // Register the add‑in; the third parameter indicates the path is relative to the workbook.
            int addInId = workbook.Worksheets.RegisterAddInFunction(addInFilePath, functionName, false);
            Console.WriteLine($"Add‑in registered with ID {addInId}");

            // 3. Save the customized workbook.
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook has been customized and saved to '{outputPath}'.");
        }
    }
}