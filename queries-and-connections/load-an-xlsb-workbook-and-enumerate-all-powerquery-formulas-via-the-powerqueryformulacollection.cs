// Title: Read an XLSB workbook and list all PowerQuery formulas from each worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to list PowerQuery formulas in an XLSB workbook with Aspose.Cells C# | C# enumerate PowerQueryFormulaCollection from worksheets using Aspose.Cells | using reflection to access PowerQuery properties in an Aspose.Cells workbook | retrieve PowerQuery queries from an XLSB file in .NET | example code to extract PowerQuery formulas from Excel binary files
// Tags: enumerate PowerQueryFormulaCollection in XLSB | read PowerQuery formulas with Aspose.Cells C# | reflection access PowerQuery properties Aspose.Cells | list worksheet PowerQuery queries .NET | load XLSB workbook using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the XLSB file exists, loads it into an Aspose.Cells Workbook, iterates through each worksheet, uses reflection to obtain the PowerQueryFormulaCollection when available, and prints the worksheet name alongside each PowerQuery formula. Missing properties are safely skipped and any runtime errors are caught and reported.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use reflection to obtain the PowerQueryFormulaCollection (may not exist in older versions)
                var pqProp = sheet.GetType().GetProperty("PowerQueryFormulaCollection");
                if (pqProp == null)
                {
                    // Property not available – skip PowerQuery processing for this worksheet
                    continue;
                }

                var pqCollection = pqProp.GetValue(sheet) as System.Collections.IEnumerable;
                if (pqCollection == null)
                {
                    continue;
                }

                // Enumerate all PowerQuery formulas in the collection using reflection
                foreach (var pqFormula in pqCollection)
                {
                    if (pqFormula == null) continue;

                    var formulaProp = pqFormula.GetType().GetProperty("Formula");
                    var formulaValue = formulaProp?.GetValue(pqFormula) as string ?? string.Empty;

                    Console.WriteLine($"Worksheet: {sheet.Name}, PowerQuery Formula: {formulaValue}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
