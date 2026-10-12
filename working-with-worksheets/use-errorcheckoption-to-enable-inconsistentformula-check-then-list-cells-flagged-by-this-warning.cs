// Title: How to enable Inconsistent Formula error checking in Aspose.Cells for .NET and retrieve the flagged cell addresses
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, activates the InconsistentFormula error‑check option, recalculates all formulas, and prints each cell address that is flagged as inconsistent. | Show how to use reflection to set Settings.ErrorCheckOptions.InconsistentFormula to true and to invoke Cells.GetInconsistentFormulaCells for compatibility with older Aspose.Cells releases. | Modify the sample so that the list of inconsistent formula cells is written to a CSV file instead of being written to the console.
// Common Searches: Aspose.Cells C# enable inconsistent formula warning and get cell list | how to retrieve cells with inconsistent formulas using Aspose.Cells .NET | reflection access ErrorCheckOptions InconsistentFormula property Aspose.Cells | list inconsistent formula cells after recalculating workbook Aspose.Cells | C# Aspose.Cells get inconsistent formula cells without compile‑time API
// Tags: Aspose.Cells inconsistent formula validation | .NET retrieve cells flagged by InconsistentFormula | reflection set Aspose.Cells ErrorCheckOptions | export inconsistent formula addresses to CSV | recalculate workbook before formula consistency analysis

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, uses reflection to turn on the InconsistentFormula error‑check option, recalculates formulas, and then calls the internal GetInconsistentFormulaCells method (also via reflection) to obtain any cells flagged for formula inconsistency. It outputs the cell addresses to the console and includes error handling for missing files and runtime exceptions.
class InconsistentFormulaChecker
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Enable the Inconsistent Formula error check option if the API is available
            var settings = workbook.Settings;
            var errorCheckProp = settings.GetType().GetProperty("ErrorCheckOptions");
            if (errorCheckProp != null)
            {
                var errorCheckOptions = errorCheckProp.GetValue(settings);
                var incFormulaProp = errorCheckOptions?.GetType().GetProperty("InconsistentFormula");
                if (incFormulaProp != null && incFormulaProp.CanWrite)
                {
                    incFormulaProp.SetValue(errorCheckOptions, true);
                }
            }

            // Recalculate formulas to ensure error checking is up‑to‑date
            workbook.CalculateFormula();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Retrieve cells flagged with the Inconsistent Formula warning (using reflection for compatibility)
            Cell[] inconsistentCells = null;
            var method = typeof(Cells).GetMethod("GetInconsistentFormulaCells", Type.EmptyTypes);
            if (method != null)
            {
                inconsistentCells = (Cell[])method.Invoke(cells, null);
            }

            // Output the results
            if (inconsistentCells != null && inconsistentCells.Length > 0)
            {
                Console.WriteLine("Cells flagged with Inconsistent Formula warning:");
                foreach (Cell cell in inconsistentCells)
                {
                    Console.WriteLine(cell.Name);
                }
            }
            else
            {
                Console.WriteLine("No inconsistent formula cells found.");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
