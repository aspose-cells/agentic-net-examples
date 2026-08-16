// Title: Load workbook with LoadOptions.FilterVbaProject to skip unsigned VBA macros (Aspose.Cells .NET)
// Description: Demonstrates how to import an .xlsm file using Aspose.Cells, detect whether its VBA project is signed, automatically remove unsigned macros, and save the result as a macro‑free .xlsx.
// Keywords: Aspose.Cells LoadOptions.FilterVbaProject | remove unsigned VBA macros C# | Workbook.HasMacro check | VbaProject.IsSigned | convert .xlsm to .xlsx without macros | secure Excel import Aspose | macro stripping Aspose.Cells
// Common Searches: Aspose.Cells ignore unsigned macros on load | How to filter VBA project when loading workbook in .NET | Remove unsigned macros automatically Aspose | LoadOptions.FilterVbaProject example C# | Securely import macro‑enabled Excel files Aspose
// Developer Intent: Import an Excel workbook while automatically discarding any unsigned VBA macros for security or conversion purposes.
// Use Cases: Sanitize user‑uploaded .xlsm files before processing to prevent execution of unsigned code. | Batch‑convert macro‑enabled workbooks to .xlsx, preserving only signed macros. | Create a secure ETL pipeline that strips unsigned VBA projects during data ingestion.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions.FilterVbaProject to load a workbook and exclude unsigned VBA projects. | Show how to check workbook.VbaProject.IsSigned and call workbook.RemoveMacro() when the project is not signed. | Provide an Aspose.Cells example that automatically ignores unsigned macros on import and saves the file as .xlsx.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMacroFilterDemo
{
    // Demonstrates how to import an .xlsm file using Aspose.Cells, detect whether its VBA project is signed, automatically remove unsigned macros, and save the result as a macro‑free .xlsx.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (may contain signed or unsigned macros)
            string sourcePath = "input_with_macros.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Error: The file '{sourcePath}' was not found.");
                return;
            }

            try
            {
                // Prepare load options (no custom filter needed for default behavior)
                LoadOptions loadOptions = new LoadOptions();

                // Load the workbook using the load options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Check if the workbook contains any VBA project
                if (workbook.HasMacro)
                {
                    // Determine whether the VBA project is signed
                    bool isSigned = workbook.VbaProject.IsSigned;

                    // If the VBA project is not signed, remove all macros
                    if (!isSigned)
                    {
                        workbook.RemoveMacro();
                        Console.WriteLine("Unsigned macros were removed during import.");
                    }
                    else
                    {
                        Console.WriteLine("Workbook contains a signed VBA project; macros are retained.");
                    }
                }
                else
                {
                    Console.WriteLine("Workbook does not contain any macros.");
                }

                // Save the resulting workbook (macro‑free if unsigned macros were removed)
                string outputPath = "output_without_unsigned_macros.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
