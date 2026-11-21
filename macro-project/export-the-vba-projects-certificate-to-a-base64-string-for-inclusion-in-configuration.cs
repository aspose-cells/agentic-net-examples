// Title: Export a VBA project's certificate to a Base64 string using Aspose.Cells for .NET
// Common Searches: C# export VBA project signature from .xlsm file to Base64 with Aspose.Cells | How to retrieve a macro-enabled workbook's VBA certificate as a Base64 string in .NET | Aspose.Cells VbaProject signature extraction to Base64 example | Convert VBA project signature to Base64 using reflection in C# | Get VBA project certificate from Excel macro file for configuration
// Tags: aspocells vba project signature API | c# base64 encoding of vba certificate | excel macro workbook certificate extraction | reflection based ExportToStream usage | load vba project from xlsm with Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject type

// The VbaCertificateExporter class loads a macro‑enabled Excel workbook, accesses its VbaProject, uses reflection to obtain the Signature object, invokes ExportToStream to write the certificate into a memory stream, and returns the certificate as a Base64‑encoded string. It includes robust error handling for missing files, absent VBA projects, and unavailable API features.
class VbaCertificateExporter
{
    // Exports the VBA project's signature from an Excel file to a Base64 string.
    public static string ExportCertificateToBase64(string excelFilePath)
    {
        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(excelFilePath))
            throw new FileNotFoundException($"The file '{excelFilePath}' was not found.");

        try
        {
            // Load the workbook (must be a macro-enabled file, e.g., .xlsm or .xlsb)
            Workbook workbook = new Workbook(excelFilePath);

            // Access the VBA project; it may be null if the workbook has no VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
                throw new InvalidOperationException("The workbook does not contain a VBA project.");

            // Use reflection to obtain the Signature property (may not exist in older versions)
            PropertyInfo signatureProp = vbaProject.GetType().GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
            if (signatureProp == null)
                throw new NotSupportedException("VBA signature functionality is not available in the current Aspose.Cells version.");

            object signature = signatureProp.GetValue(vbaProject);
            if (signature == null)
                throw new InvalidOperationException("The VBA project does not have an associated signature.");

            // Use reflection to call ExportToStream on the signature object
            MethodInfo exportMethod = signature.GetType().GetMethod("ExportToStream", new[] { typeof(Stream) });
            if (exportMethod == null)
                throw new NotSupportedException("ExportToStream method is not available on the VBA signature object.");

            using (MemoryStream ms = new MemoryStream())
            {
                exportMethod.Invoke(signature, new object[] { ms });
                return Convert.ToBase64String(ms.ToArray());
            }
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context while preserving stack trace
            throw new ApplicationException("Failed to export VBA signature.", ex);
        }
    }

    // Example usage
    static void Main()
    {
        string excelPath = "SampleWithVba.xlsm";

        try
        {
            string base64Cert = ExportCertificateToBase64(excelPath);
            Console.WriteLine("Signature Base64:");
            Console.WriteLine(base64Cert);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
