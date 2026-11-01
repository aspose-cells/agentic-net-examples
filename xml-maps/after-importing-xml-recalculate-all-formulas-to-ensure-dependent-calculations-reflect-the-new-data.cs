// Title: Recalculate all workbook formulas after loading an XML spreadsheet with Aspose.Cells for .NET (C#)
// Common Searches: how to force formula recalculation after loading an XML file with Aspose.Cells in C# | Aspose.Cells C# recalculate workbook formulas after importing XML data | C# load XML spreadsheet and evaluate formulas using Aspose.Cells | using LoadOptions to import XML and update dependent calculations in Aspose.Cells | convert XML to XLSX while preserving and recalculating formulas with Aspose.Cells
// Tags: Aspose.Cells Workbook.CalculateFormula after XML import | C# LoadOptions LoadFormat.Xml for spreadsheet import | XML to XLSX conversion with formula evaluation Aspose.Cells | update dependent calculations in Aspose.Cells workbook | recalculate formulas programmatically Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input XML file, loads it into an Aspose.Cells Workbook using LoadOptions set to XML format, forces a full formula recalculation with CalculateFormula, and then saves the workbook as an XLSX file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xml";
            const string outputPath = "output.xlsx";

            // Ensure the input XML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XML spreadsheet with appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xml);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate all formulas after loading data
            workbook.CalculateFormula();

            // Save the workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
