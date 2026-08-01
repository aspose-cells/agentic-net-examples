// Title: Use LoadOptions.FilterVbaProject to load an .xlsm file without unsigned macros and save as macro‑free .xlsx (Aspose.Cells C#)
// Description: Demonstrates how to create a LoadOptions object with FilterVbaProject enabled, load a macro‑enabled workbook while ignoring unsigned VBA projects, verify the HasMacro flag, and export the file to XLSX where all macros are stripped. Includes file‑existence validation and error handling.
// Keywords: Aspose.Cells LoadOptions.FilterVbaProject | exclude unsigned macros C# | load .xlsm without VBA | convert .xlsm to .xlsx Aspose.Cells | remove macros on import | macro‑free Excel export | HasMacro property
// Common Searches: Aspose.Cells ignore unsigned VBA projects | LoadOptions.FilterVbaProject example | C# load .xlsm without macros | convert macro enabled workbook to macro free | remove VBA macros when saving to XLSX
// Developer Intent: Load a macro‑enabled Excel file while filtering out unsigned VBA code and save it as a macro‑free XLSX document using Aspose.Cells for .NET.
// Use Cases: Distribute a workbook to users who must not receive any VBA code. | Pre‑process incoming .xlsm files in a secure pipeline by stripping unsigned macros. | Validate whether an imported workbook contains macros before applying further business logic.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions.FilterVbaProject to load an .xlsm file, skip unsigned VBA projects, and save the result as .xlsx. | Explain the effect of the FilterVbaProject option on macro handling during workbook import with Aspose.Cells. | Provide a step‑by‑step tutorial for checking file existence, loading a macro‑enabled workbook with filtered VBA, inspecting the HasMacro property, and exporting a clean XLSX file.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a LoadOptions object with FilterVbaProject enabled, load a macro‑enabled workbook while ignoring unsigned VBA projects, verify the HasMacro flag, and export the file to XLSX where all macros are stripped. Includes file‑existence validation and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "sample_with_macro.xlsm";
        const string outputPath = "output_without_unsigned_macros.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (macros will be ignored when saving to XLSX)
            Workbook workbook = new Workbook(inputPath);

            // Display whether the loaded workbook contains macros
            Console.WriteLine("HasMacro after load: " + workbook.HasMacro);

            // Save as XLSX; macros are automatically removed because XLSX does not support them
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
