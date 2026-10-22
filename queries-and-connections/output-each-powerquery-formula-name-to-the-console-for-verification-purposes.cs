// Title: List Power Query formula names from each worksheet in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells and prints the names of all Power Query formulas defined in each worksheet. | Show how to modify the sample to also display the worksheet name together with each Power Query formula name in the console output. | Create a reusable method that returns a dictionary mapping worksheet names to their Power Query formula names by using reflection on Aspose.Cells objects.
// Common Searches: Aspose.Cells C# read Power Query names from workbook | how to enumerate PowerQueryFormulaCollection in Aspose.Cells | retrieve Power Query query names from Excel using Aspose.Cells .NET | list all Power Query queries per sheet with Aspose.Cells reflection
// Tags: Aspose.Cells PowerQueryFormulaCollection enumeration | C# extract Power Query names from Excel workbook | reflection access PowerQuery formulas Aspose.Cells | list Power Query queries per worksheet .NET | output Power Query formula names console

using System;
using System.IO;
using Aspose.Cells;
using System.Collections;

// The example verifies the presence of an input.xlsx file, loads it with Aspose.Cells, iterates through each worksheet, uses reflection to obtain the PowerQueryFormulaCollection when available, and writes each Power Query formula's Name to the console while handling missing properties and exceptions gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: File '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use reflection to obtain the PowerQueryFormulaCollection (if supported by the library version)
                var pqCollectionProp = sheet.GetType().GetProperty("PowerQueryFormulaCollection");
                if (pqCollectionProp == null)
                {
                    // Property not available in this version; skip to next sheet
                    continue;
                }

                var pqCollection = pqCollectionProp.GetValue(sheet, null) as IEnumerable;
                if (pqCollection == null)
                {
                    continue;
                }

                // Output each PowerQuery formula name to the console
                foreach (var pqFormula in pqCollection)
                {
                    var nameProp = pqFormula.GetType().GetProperty("Name");
                    var name = nameProp?.GetValue(pqFormula, null);
                    Console.WriteLine(name?.ToString() ?? "Unnamed Formula");
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
