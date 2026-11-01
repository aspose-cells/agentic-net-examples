// Title: How to load an XLSB workbook and list all Power Query formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a binary Excel (.xlsb) file with Aspose.Cells and prints each Power Query formula to the console. | Show a robust C# example that verifies the PowerQueryFormulaCollection exists via reflection and safely iterates over its items. | Provide a C# snippet that detects missing Power Query support in Aspose.Cells, logs an appropriate warning, and then enumerates any available formulas.
// Common Searches: C# Aspose.Cells retrieve Power Query formulas from an .xlsb workbook | How to enumerate PowerQueryFormulaCollection in Aspose.Cells using reflection | List all Power Query queries in a binary Excel file with Aspose.Cells .NET | Check if Power Query support is available in Aspose.Cells before accessing formulas
// Tags: open binary Excel workbook using Aspose.Cells | enumerate PowerQueryFormulaCollection in C# | access Power Query formulas via reflection Aspose.Cells | detect Power Query support in Aspose.Cells workbook | output Power Query formulas to console

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks that the specified .xlsb file exists, loads it into an Aspose.Cells Workbook, uses reflection to obtain the Worksheets.PowerQueryFormulaCollection when available, iterates through each item to read its Formula property, prints the formulas to the console, and gracefully handles cases where Power Query support is absent or errors occur.
    class Program
    {
        static void Main()
        {
            const string filePath = "input.xlsb";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the XLSB workbook
                Workbook workbook = new Workbook(filePath);

                // Attempt to retrieve Power Query formulas via reflection (compatible with multiple versions)
                var worksheets = workbook.Worksheets;
                var pqProp = worksheets.GetType().GetProperty("PowerQueryFormulaCollection");
                if (pqProp == null)
                {
                    Console.WriteLine("Power Query formulas are not supported in this version of Aspose.Cells.");
                    return;
                }

                var pqCollection = pqProp.GetValue(worksheets) as IEnumerable;
                if (pqCollection == null)
                {
                    Console.WriteLine("No Power Query formulas found.");
                    return;
                }

                // Enumerate and display each Power Query formula
                foreach (var pqItem in pqCollection)
                {
                    var formulaProp = pqItem.GetType().GetProperty("Formula");
                    var formula = formulaProp?.GetValue(pqItem) as string;
                    Console.WriteLine(formula);
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
