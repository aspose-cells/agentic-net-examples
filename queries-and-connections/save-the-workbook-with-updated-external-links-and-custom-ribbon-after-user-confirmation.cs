using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the main workbook that contains external links
        Workbook mainWorkbook = new Workbook("Main.xlsx");

        // Load external workbook(s) that provide the latest data for the links
        Workbook externalWorkbook = new Workbook("External.xlsx");

        // Update external links in the main workbook with data from the external workbook(s)
        mainWorkbook.UpdateLinkedDataSource(new Workbook[] { externalWorkbook });

        // Recalculate formulas after updating links (optional but often needed)
        mainWorkbook.CalculateFormula();

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
        mainWorkbook.RibbonXml = ribbonXml;

        // Ask the user for confirmation before saving
        Console.Write("Save the workbook with updated links and custom ribbon? (y/n): ");
        string answer = Console.ReadLine();

        if (!string.IsNullOrEmpty(answer) && answer.Trim().ToLower() == "y")
        {
            // Save the workbook (XLSM is required for Ribbon customizations)
            mainWorkbook.Save("Main_Updated.xlsm", SaveFormat.Xlsm);
            Console.WriteLine("Workbook saved successfully.");
        }
        else
        {
            Console.WriteLine("Save operation cancelled by the user.");
        }

        // Release resources
        mainWorkbook.Dispose();
        externalWorkbook.Dispose();
    }
}