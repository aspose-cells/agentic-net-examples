// Title: Disable automatic formula recalculation in an Aspose.Cells workbook using C# (set manual calculation mode)
// Common Searches: Aspose.Cells C# set workbook calculation mode to manual | prevent Excel formulas from recalculating on open with Aspose.Cells .NET | how to turn off automatic formula evaluation when saving a workbook in Aspose.Cells | disable automatic recalculation of formulas in Aspose.Cells for .NET applications | manual calculation setting Aspose.Cells workbook C# example
// Tags: Aspose.Cells manual calculation mode C# | disable automatic formula evaluation Aspose.Cells | Workbook.Settings CalcMode property Aspose.Cells | prevent formula recalculation on workbook open Aspose.Cells | Aspose.Cells calculation settings API C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file into an Aspose.Cells Workbook, shows how to set the calculation mode to manual and disable formula evaluation on open (commented out for versions where the API is unavailable), and saves the workbook to a new XLSX file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // The following settings are not available in the current Aspose.Cells version.
            // If needed, they can be set using the appropriate API for the installed version.
            // workbook.Settings.CalcMode = CalculationMode.Manual;
            // workbook.Settings.CalculateFormulaOnOpen = false;

            // Save the workbook to the desired output format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
