// Title: How to load an XLSB workbook and enumerate its Power Query formulas with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an XLSB file with Aspose.Cells and iterates through workbook.DataMashup.PowerQueryFormulas to display each formula's name, definition, type, group name, and description. | Provide a C# snippet that writes the names and definitions of all Power Query formulas from a loaded XLSB workbook to a CSV file using Aspose.Cells. | Show how to check for the existence of Power Query formulas in a workbook and handle the case when none are found, using Aspose.Cells in .NET.
// Common Searches: C# Aspose.Cells how to list Power Query formulas from an XLSB file | retrieve Power Query mashup data programmatically with Aspose.Cells .NET | enumerate PowerQueryFormulaCollection in a binary Excel workbook using Aspose.Cells | read Power Query formula definitions from XLSB using Aspose.Cells API | Aspose.Cells DataMashup PowerQueryFormulas example in C#
// Tags: Aspose.Cells read PowerQueryFormulaCollection XLSB | enumerate Power Query formulas .NET | extract DataMashup PowerQueryFormulas C# | list Power Query formula definitions Aspose.Cells | access workbook.DataMashup PowerQueryFormulas

using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

namespace AsposeCellsPowerQueryDemo
{
    // The example loads an XLSB workbook with Aspose.Cells, accesses its DataMashup, verifies that the PowerQueryFormulas collection is present, and iterates through each PowerQueryFormula to print its name, definition, type, group name, and description.
    class Program
    {
        static void Main()
        {
            // Path to the XLSB workbook that contains Power Query formulas
            string sourcePath = "source.xlsb";

            // Load the workbook (XLSB format is automatically detected)
            Workbook workbook = new Workbook(sourcePath);

            // Access the mashup data of the workbook
            DataMashup mashup = workbook.DataMashup;

            // Ensure the mashup and its PowerQueryFormulas collection are available
            if (mashup != null && mashup.PowerQueryFormulas != null && mashup.PowerQueryFormulas.Count > 0)
            {
                Console.WriteLine($"Found {mashup.PowerQueryFormulas.Count} Power Query formula(s):");

                // Enumerate each PowerQueryFormula in the collection
                foreach (PowerQueryFormula formula in mashup.PowerQueryFormulas)
                {
                    Console.WriteLine("--------------------------------------------------");
                    Console.WriteLine($"Formula Name       : {formula.Name}");
                    Console.WriteLine($"Formula Definition : {formula.FormulaDefinition}");
                    Console.WriteLine($"Formula Type       : {formula.Type}");
                    Console.WriteLine($"Group Name         : {formula.GroupName}");
                    Console.WriteLine($"Description        : {formula.Description}");
                }
            }
            else
            {
                Console.WriteLine("No Power Query formulas found in the workbook.");
            }

            // (Optional) Save the workbook to a new file if any modifications were made
            // workbook.Save("output.xlsb");
        }
    }
}
