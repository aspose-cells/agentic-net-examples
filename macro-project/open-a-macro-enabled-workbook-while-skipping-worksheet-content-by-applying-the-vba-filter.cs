using System;
using System.IO;
using Aspose.Cells;

class OpenMacroWorkbook
{
    static void Main()
    {
        // Path to the macro‑enabled workbook
        string filePath = "sample_with_macro.xlsm";

        try
        {
            // Prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
                return;
            }

            // Load only workbook structure and VBA, skip cell data
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new LoadFilter(LoadDataFilterOptions.Structure | LoadDataFilterOptions.VBA)
            };

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Output whether the workbook contains macros
            Console.WriteLine("HasMacro: " + workbook.HasMacro);
        }
        catch (Exception ex)
        {
            // Runtime safety
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}