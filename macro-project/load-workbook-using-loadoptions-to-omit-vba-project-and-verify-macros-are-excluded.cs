using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroOmitDemo
{
    // Custom load filter that loads everything except VBA projects
    class NoVbaLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load all data but exclude VBA (bitwise remove VBA flag)
            LoadDataFilterOptions = LoadDataFilterOptions.All & ~LoadDataFilterOptions.VBA;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a macro‑enabled workbook (e.g., .xlsm) that contains VBA code
            string inputPath = "sample_with_macro.xlsm";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Configure load options to use the custom filter that skips VBA
                LoadOptions loadOptions = new LoadOptions
                {
                    LoadFilter = new NoVbaLoadFilter()
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Verify that macros have been omitted
                Console.WriteLine("HasMacro after load with filter: " + workbook.HasMacro); // Expected: False

                // Optionally save the macro‑free workbook
                string outputPath = "macro_free.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Macro‑free workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}