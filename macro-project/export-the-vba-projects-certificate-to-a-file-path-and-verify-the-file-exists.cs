// Title: Export a VBA project's digital certificate to a .cer file and verify its existence with Aspose.Cells for .NET (C#)
// Common Searches: how to export a VBA project's certificate to a .cer file using Aspose.Cells in C# | C# code to check if an Excel workbook contains a VBA project before exporting its certificate | using reflection to call ExportCertificate on Aspose.Cells VbaProject | verify that a certificate file was created after exporting from a macro-enabled workbook in .NET | Aspose.Cells export VBA project certificate and confirm file creation
// Tags: Aspose.Cells VbaProject ExportCertificate | C# reflection invoke Aspose.Cells method | validate exported certificate file existence | detect VBA project in .xlsm workbook | export certificate to .cer using Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject extension methods (if available)

// The example loads a macro-enabled .xlsm workbook, ensures it contains a VBA project, uses reflection to invoke the VbaProject.ExportCertificate method to write the project's digital certificate to a specified .cer file, and then checks that the exported certificate file exists, handling any errors that may occur.
class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the Excel file containing the VBA project
        string workbookPath = @"C:\Path\To\YourWorkbook.xlsm";

        // Path where the certificate will be exported
        string certificatePath = @"C:\Path\To\ExportedCertificate.cer";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the workbook actually contains a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Try to invoke ExportCertificate via reflection (covers different API versions)
            MethodInfo exportMethod = workbook.VbaProject.GetType()
                .GetMethod("ExportCertificate", new[] { typeof(string) });

            if (exportMethod != null)
            {
                // Export the certificate
                exportMethod.Invoke(workbook.VbaProject, new object[] { certificatePath });
                Console.WriteLine("Certificate export method invoked successfully.");
            }
            else
            {
                Console.WriteLine("ExportCertificate method is not available in this Aspose.Cells version.");
                return;
            }

            // Verify that the certificate file now exists
            bool fileExists = File.Exists(certificatePath);
            Console.WriteLine($"Certificate export successful: {fileExists}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
