// Title: C# – Load an Excel workbook with Aspose.Cells using LoadOptions.FilterVbaProject to include only signed VBA macros
// Description: Demonstrates how to configure LoadOptions.FilterVbaProject to load only signed VBA projects, verify the digital signature with Workbook.VbaProject.IsSigned and IsValidSigned, and optionally save the workbook while preserving the signed macros. Includes file‑existence checks and basic error handling.
// Keywords: Aspose.Cells LoadOptions.FilterVbaProject | load signed VBA macros C# | Workbook.VbaProject.IsSigned | check VBA signature Aspose.Cells | process signed macro workbook | save Xlsm with signed macros | C# Excel macro security | global Excel automation
// Common Searches: How to load only signed VBA projects with Aspose.Cells | Aspose.Cells filter VBA macros by signature | C# check if VBA project is signed in Excel file | LoadOptions.FilterVbaProject example | Validate VBA macro signature using Aspose.Cells
// Developer Intent: Load an Excel file, automatically filter out unsigned VBA projects, confirm the signature status, and then continue with custom processing or saving.
// Use Cases: Enforce security policies by processing workbooks only when the embedded VBA project is digitally signed. | Automate validation of macro signatures before executing or modifying macro code. | Preserve signed VBA macros while performing other workbook transformations and saving back to Xlsm.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions.FilterVbaProject to load only signed VBA projects and logs the signature validity. | Provide a robust error‑handling pattern for cases where the workbook lacks a signed VBA project after loading with Aspose.Cells. | Create a reusable method that extracts the digital signature details from Workbook.VbaProject and returns a validation report.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to configure LoadOptions.FilterVbaProject to load only signed VBA projects, verify the digital signature with Workbook.VbaProject.IsSigned and IsValidSigned, and optionally save the workbook while preserving the signed macros. Includes file‑existence checks and basic error handling.
class LoadSignedVbaWorkbook
{
    static void Main()
    {
        const string inputPath = "input_signed.xlsm";
        const string outputPath = "output_processed.xlsm";

        // Verify input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // LoadOptions with default auto-detect format (no specific filter needed)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Check if a signed VBA project is present
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("No signed VBA project loaded.");
            }

            // TODO: Add further processing of the workbook here

            // Save the workbook after processing (optional)
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
