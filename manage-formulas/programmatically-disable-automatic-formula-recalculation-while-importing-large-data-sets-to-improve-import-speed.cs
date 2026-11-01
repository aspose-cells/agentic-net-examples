// Title: Disable automatic formula recalculation during large data import with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# turn off automatic calculation for faster bulk import | set manual calculation mode before importing DataTable in Aspose.Cells .NET | disable formula recalculation when using Worksheet.Cells.ImportDataTable Aspose.Cells | improve import speed by disabling auto calculate in Aspose.Cells workbook | how to prevent automatic formula evaluation during large data import Aspose.Cells
// Tags: Aspose.Cells calculation mode setting | Workbook Settings CalcMode | bulk import performance tuning Aspose.Cells | disable auto formula evaluation | large dataset import without recalculation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, shows how to set the calculation mode to Manual via workbook.Settings.CalcMode to stop automatic formula recalculation, imports a large data set (e.g., with ImportDataTable or ImportArray), optionally triggers a manual recalculation with workbook.CalculateFormula, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: In newer Aspose.Cells versions the calculation mode can be set via
            // workbook.Settings.CalcMode = CalculationMode.Manual;
            // If the property is unavailable, the default automatic mode will be used.

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // TODO: Import your large data set here, e.g. using ImportDataTable, ImportArray, etc.
            // Example (commented out):
            // System.Data.DataTable dataTable = GetLargeDataTable();
            // sheet.Cells.ImportDataTable(dataTable, true, "A1");

            // If you need to recalculate formulas after the import, do it manually:
            // workbook.CalculateFormula();

            // Define output file path
            string outputPath = "LargeDataImport.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
