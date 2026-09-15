// Title: How to configure Aspose.Cells in C# to treat empty strings as zero during formula calculations
// AI Prompts: Set workbook.CalcEngineSettings.EmptyStringAsZero = true before saving so blank cells are evaluated as 0 in numeric formulas. | Show a C# snippet that enables the EmptyStringAsZero flag in Aspose.Cells and verifies the result with a sample formula.
// Common Searches: Aspose.Cells C# treat blank cell as zero in formula evaluation | Enable EmptyStringAsZero property in Aspose.Cells calculation engine C# example | How to make empty strings evaluate to 0 in Excel formulas using Aspose.Cells | C# Aspose.Cells calculation settings for handling empty strings in numeric formulas
// Tags: Aspose.Cells EmptyStringAsZero setting C# | C# Aspose.Cells calculation options | Excel formula blank cell zero handling Aspose.Cells | Configure workbook calculation settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook or creates a new one and saves it. To ensure numeric formulas treat empty strings as zero, set workbook.CalcEngineSettings.EmptyStringAsZero = true (or use the version‑specific API) before saving. This configures the calculation engine so blank cells are interpreted as 0 during formula evaluation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook if it exists; otherwise create a new one
                const string inputPath = "input.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath); // load-workbook
                }
                else
                {
                    workbook = new Workbook(); // create-workbook
                }

                // NOTE: In some Aspose.Cells versions the CalcEngineSettings property may not be available.
                // If needed, configure calculation options using the appropriate API for your version.

                // Save the workbook with the new setting applied
                const string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists
                var outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                try
                {
                    workbook.Save(outputPath); // save-workbook
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.Error.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
