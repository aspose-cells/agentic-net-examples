using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the main workbook (replace with your actual file path)
        string mainPath = "input.xlsx";
        Workbook workbook = new Workbook(mainPath);

        // Load an external workbook that provides the latest data for linked sources
        // (replace with the actual external file path)
        Workbook externalWorkbook = new Workbook("ExternalData.xlsx");

        // Update all external links in the main workbook using the external workbook
        workbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

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

        // Apply the custom Ribbon UI to the workbook
        workbook.RibbonXml = ribbonXml;

        // Ask the user whether to save the changes
        Console.Write("Save the workbook with updated links and custom ribbon? (y/n): ");
        string response = Console.ReadLine();

        if (!string.IsNullOrEmpty(response) && response.Trim().Equals("y", StringComparison.OrdinalIgnoreCase))
        {
            // Save the workbook (using .xlsm to support Ribbon UI)
            workbook.Save("output.xlsm");
            Console.WriteLine("Workbook saved successfully.");
        }
        else
        {
            Console.WriteLine("Save operation cancelled by the user.");
        }

        // Clean up resources
        workbook.Dispose();
        externalWorkbook.Dispose();
    }
}