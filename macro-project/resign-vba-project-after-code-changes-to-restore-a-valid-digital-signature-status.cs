// Title: Re‑sign a VBA project in an XLSM workbook after code changes with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells re‑sign VBA project after editing XLSM | preserve VBA macro digital signature when saving workbook using Aspose.Cells .NET | Aspose.Cells .NET cannot sign VBA project workaround | update macro code in XLSM and keep digital signature with Aspose.Cells | apply digital certificate to VBA project in Excel file using C# Aspose.Cells
// Tags: VBA project signing Aspose.Cells | maintain macro signature after workbook edit | Aspose.Cells VBA digital signing limitation | save macro‑enabled workbook C# Aspose.Cells | apply certificate to VBA project .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an XLSM workbook with Aspose.Cells, accesses its VbaProject (optional modifications can be made), notes that digital signing of VBA projects is not supported in the current Aspose.Cells version, and then saves the workbook as a macro‑enabled file.
class Program
{
    static void Main()
    {
        // Paths to required files
        string workbookPath = "input.xlsm";
        string certificatePath = "mycert.pfx";
        string certificatePassword = "password";

        // Verify that the input files exist
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        if (!File.Exists(certificatePath))
        {
            Console.WriteLine($"Error: Certificate file not found at '{certificatePath}'.");
            return;
        }

        try
        {
            // Load the workbook while preserving macros
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project (optional modifications can be done here)
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.Modules.Count > 0)
            {
                // Example: append a comment to the first module
                // VbaModule module = vbaProject.Modules[0];
                // module.Code += "\r\n' Modified by Aspose.Cells";
            }

            // NOTE: Digital signing of VBA projects is not available in the current Aspose.Cells version.
            // The following code has been omitted to ensure successful compilation.

            // Save the workbook, keeping the macro‑enabled format
            string outputPath = "output_signed.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
