// Title: Export a VBA project's digital certificate to a .cer file using Aspose.Cells for .NET (C#)
// Common Searches: how to export VBA project certificate to a .cer file using Aspose.Cells C# | Aspose.Cells extract VBA macro certificate to disk in .NET | C# code to save VBA project digital certificate with Aspose.Cells | using reflection to export VBA certificate from an Excel workbook Aspose.Cells | export VBA project certificate from .xlsx using Aspose.Cells library
// Tags: Aspose.Cells VBA certificate export | C# extract VBA macro certificate | reflection access VBA certificate property | save certificate as .cer file .NET | export VBA project digital certificate

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies the presence of a VBA project, uses reflection to obtain the project's Certificate object, and invokes its Export method to write the certificate to a .cer file, handling missing files or unsupported API versions gracefully.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string certPath = "VbaCertificate.cer";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains a VBA project with a certificate
            Workbook workbook = new Workbook(inputPath);

            // Ensure the workbook has a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Try to obtain the Certificate property via reflection (covers different API versions)
            var vbaProject = workbook.VbaProject;
            var certProp = vbaProject.GetType().GetProperty("Certificate");
            if (certProp == null)
            {
                Console.WriteLine("The VBA project does not expose a Certificate property in this Aspose.Cells version.");
                return;
            }

            var certificate = certProp.GetValue(vbaProject);
            if (certificate == null)
            {
                Console.WriteLine("The VBA project does not contain a certificate.");
                return;
            }

            // Use reflection to call the Export method on the certificate object
            var exportMethod = certificate.GetType().GetMethod("Export", new[] { typeof(string) });
            if (exportMethod == null)
            {
                Console.WriteLine("The Certificate object does not have an Export method.");
                return;
            }

            exportMethod.Invoke(certificate, new object[] { certPath });
            Console.WriteLine($"Certificate exported successfully to {certPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
