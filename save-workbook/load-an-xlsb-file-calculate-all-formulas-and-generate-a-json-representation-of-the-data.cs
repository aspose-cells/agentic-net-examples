// Title: Load an XLSB workbook, evaluate all formulas, and save the result as JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a binary Excel (XLSB) file with Aspose.Cells, forces full formula recalculation, and writes the workbook contents to a JSON file. | Show how to configure LoadOptions for XLSB, call Workbook.CalculateFormula, and use SaveFormat.Json to export the data. | Demonstrate converting an XLSB workbook to JSON after evaluating formulas, including sample file paths and console output.
// Common Searches: Aspose.Cells C# load XLSB file and export to JSON after calculating formulas | How to recalculate formulas in an XLSB workbook using Aspose.Cells before saving as JSON | C# example for converting binary Excel (XLSB) to JSON with formula evaluation using Aspose.Cells | SaveFormat.Json usage with XLSB workbook in Aspose.Cells .NET
// Tags: load XLSB workbook Aspose.Cells | calculate workbook formulas Aspose.Cells | export workbook to JSON Aspose.Cells | SaveFormat.Json with XLSB data | C# binary Excel to JSON conversion

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an XLSB workbook using LoadOptions, recalculates all formulas with CalculateFormula, and saves the workbook data as a JSON file via SaveFormat.Json.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSB file
            string xlsbPath = "input.xlsb";

            // Load the XLSB workbook with appropriate load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
            Workbook workbook = new Workbook(xlsbPath, loadOptions);

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook data as JSON
            string jsonOutputPath = "output.json";
            workbook.Save(jsonOutputPath, SaveFormat.Json);

            Console.WriteLine($"Workbook formulas calculated and saved as JSON to '{jsonOutputPath}'.");
        }
    }
}
