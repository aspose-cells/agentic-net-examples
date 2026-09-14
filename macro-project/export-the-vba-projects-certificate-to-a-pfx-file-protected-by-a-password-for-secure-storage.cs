// Title: Export a VBA project's digital certificate to a password‑protected PFX file using Aspose.Cells for .NET (API limitation)
// AI Prompts: Write C# code with Aspose.Cells that detects a VBA project in an XLSM workbook and, if the library supports it, writes the project's signing certificate to a .pfx file encrypted with a supplied password. | Enhance the sample with robust error handling: check that the input file exists, handle the scenario where no VBA project is found, and catch any exceptions that may occur during the certificate export process. | Add a verification step after the export attempt that confirms the .pfx file was created and logs a clear success or failure message.
// Common Searches: Aspose.Cells .NET export VBA macro certificate to PFX | C# retrieve VBA project digital signature from Excel workbook | How to create a password protected PFX from a VBA project using Aspose.Cells | Missing VBA project handling in Aspose.Cells workbook processing
// Tags: VbaProject certificate export Aspose.Cells | password-protected PFX generation C# | Aspose.Cells VBA digital signature extraction | detect absent VBA project workbook | Aspose.Cells API limitation certificate export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Namespace containing VbaProject

// The example loads an XLSM workbook, checks for the presence of a VBA project, and attempts to export its digital certificate to a password‑protected .pfx file. Because the current Aspose.Cells API does not provide a method for exporting a VBA project's certificate, the code only logs a message indicating this limitation while demonstrating proper file existence checks and error handling.
class ExportVbaCertificate
{
    static void Main()
    {
        const string inputFile = "input.xlsm";
        const string outputCertFile = "VbaProjectCertificate.pfx";
        const string password = "SecurePassword123";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputFile);

            // Access the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Aspose.Cells does not expose a direct API to export a VBA project's digital certificate.
            // If such functionality becomes available, replace the following placeholder with the correct call.
            if (vbaProject != null)
            {
                Console.WriteLine("VBA project detected, but exporting its certificate is not supported by the current Aspose.Cells API.");
                // Example placeholder for future implementation:
                // vbaProject.Signature.Export(outputCertFile, password);
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
