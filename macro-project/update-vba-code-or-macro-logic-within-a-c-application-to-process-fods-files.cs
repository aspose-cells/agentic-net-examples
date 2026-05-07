using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsFodsProcessing
{
    public class FodsMacroHandler
    {
        public static void ProcessFods(string fodsPath, string outputPath)
        {
            // Load the FODS file with explicit load options for the FODS format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);
            Workbook workbook = new Workbook(fodsPath, loadOptions);

            // If the workbook contains any VBA/macros, remove them
            if (workbook.HasMacro)
            {
                workbook.RemoveMacro();
            }

            // Save the processed workbook as an XLSX file (macro‑free)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        // Example usage
        public static void Main()
        {
            string inputFods = "sample.fods";
            string outputXlsx = "sample_processed.xlsx";

            ProcessFods(inputFods, outputXlsx);

            Console.WriteLine($"FODS file processed and saved to {outputXlsx}");
        }
    }
}