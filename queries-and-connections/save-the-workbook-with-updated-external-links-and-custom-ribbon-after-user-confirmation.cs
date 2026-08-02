using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Set custom Ribbon XML
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
        workbook.RibbonXml = ribbonXml;

        // Update external links if any exist
        if (workbook.Worksheets.ExternalLinks.Count > 0)
        {
            // Example: modify the first external link's data source
            workbook.Worksheets.ExternalLinks[0].DataSource = "updated_external.xlsx";

            // Load the external workbook that provides the latest data
            Workbook externalWb = new Workbook("updated_external.xlsx");

            // Refresh linked data sources
            workbook.UpdateLinkedDataSource(new Workbook[] { externalWb });

            // Recalculate formulas to reflect updated data
            workbook.CalculateFormula();
        }

        // Ask user for confirmation before saving
        Console.Write("Save the workbook with changes? (y/n): ");
        string response = Console.ReadLine();
        if (!string.IsNullOrEmpty(response) && response.Equals("y", StringComparison.OrdinalIgnoreCase))
        {
            // Save as macro‑enabled workbook to retain Ribbon XML
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
            Console.WriteLine("Workbook saved as output.xlsm");
        }
        else
        {
            Console.WriteLine("Save operation cancelled.");
        }

        // Release resources
        workbook.Dispose();
    }
}