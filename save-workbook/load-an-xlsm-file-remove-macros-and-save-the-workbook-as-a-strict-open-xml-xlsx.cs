using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroRemoval
{
    public class MacroRemovalDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the macro-enabled source workbook (XLSM)
            string sourcePath = "input_with_macros.xlsm";

            // Path for the macro‑free strict Open XML workbook (XLSX)
            string outputPath = "output_strict.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.Error.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(sourcePath);

            // Remove all VBA/macros from the workbook
            workbook.RemoveMacro();

            // Set OOXML compliance to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Save the workbook as a strict XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Macros removed and saved as strict XLSX: {outputPath}");
        }
    }
}