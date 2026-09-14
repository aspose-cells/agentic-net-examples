// Title: Recalculate all workbook formulas after loading XML data with Aspose.Cells and save as XLSX in C#
// AI Prompts: Load an XML file into an Aspose.Cells Workbook, invoke CalculateFormula to refresh all dependent calculations, and write the result to an XLSX file using C#. | Write C# code that verifies the XML source exists, imports it with LoadOptions(LoadFormat.Xml), triggers a full formula recalculation, and saves the workbook as output.xlsx. | Generate a C# example showing how to import XML data, recalculate all formulas, and export the workbook to Excel format with Aspose.Cells.
// Common Searches: how to force formula recalculation after importing XML with Aspose.Cells .NET | Aspose.Cells C# load XML and update dependent formulas before saving | C# example for CalculateFormula after loading XML data into workbook | save workbook as XLSX after XML import and formula refresh using Aspose.Cells | recalculate all formulas programmatically in Aspose.Cells after XML map load
// Tags: import XML data with Aspose.Cells | recalculate formulas programmatically | export workbook to XLSX C# | Aspose.Cells LoadOptions XML usage | C# formula refresh after XML load

using System;
using System.IO;
using Aspose.Cells;

// The C# program checks for the presence of data.xml, loads it into an Aspose.Cells Workbook using LoadFormat.Xml, calls CalculateFormula to update all dependent calculations, and saves the refreshed workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML source file
            string xmlPath = "data.xml";

            // Verify that the XML file exists to avoid FileNotFoundException
            if (!File.Exists(xmlPath))
            {
                Console.WriteLine($"Error: The file '{xmlPath}' was not found.");
                return;
            }

            // Load XML data into a new workbook using LoadOptions
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
            Workbook workbook = new Workbook(xmlPath, loadOptions);

            // Recalculate all formulas so that dependent calculations reflect the new data
            workbook.CalculateFormula();

            // Path for the output Excel file
            string outputPath = "output.xlsx";

            // Save the updated workbook to an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
