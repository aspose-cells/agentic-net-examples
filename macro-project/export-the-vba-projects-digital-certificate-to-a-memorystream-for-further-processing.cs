// Title: Export VBA project digital certificate to a MemoryStream using Aspose.Cells for .NET
// AI Prompts: Generate a C# method that opens an .xlsm workbook with Aspose.Cells, accesses its VbaProject, and returns the embedded digital certificate as a MemoryStream. | Show how to employ reflection in C# to obtain the DigitalSignature property from a VbaProject and invoke its Export method to write the certificate into a stream. | Implement comprehensive error handling for cases where the workbook lacks a VBA project, the DigitalSignature property is missing, or the Export method cannot be called.
// Common Searches: how to extract VBA digital certificate from an xlsm file using Aspose.Cells in C# | Aspose.Cells C# export VbaProject DigitalSignature to stream | retrieve macro signing certificate from Excel workbook programmatically | C# reflection get VBA project certificate with Aspose.Cells | save VBA project digital signature to .cer file using Aspose.Cells
// Tags: export VBA digital certificate to MemoryStream Aspose.Cells | retrieve VbaProject DigitalSignature via reflection C# | load .xlsm workbook with macros Aspose.Cells | handle missing VBA project exception Aspose.Cells | save VBA certificate as .cer file C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace Example
{
    // The example loads an .xlsm workbook, accesses its VbaProject, uses reflection to obtain the DigitalSignature property, invokes the Export method to write the certificate into a MemoryStream, and returns the stream while providing detailed error handling for missing projects, unsupported versions, and export failures.
    public class VbaCertificateExporter
    {
        /// <param name="excelFilePath">Path to the .xlsm workbook containing the VBA project.</param>
        /// <returns>MemoryStream containing the exported digital certificate.</returns>
        public MemoryStream ExportVbaCertificate(string excelFilePath)
        {
            if (string.IsNullOrWhiteSpace(excelFilePath))
                throw new ArgumentException("File path is null or empty.", nameof(excelFilePath));

            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException("Excel file not found.", excelFilePath);

            try
            {
                // Load the workbook (preserves macros automatically)
                Workbook workbook = new Workbook(excelFilePath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                if (vbaProject == null)
                    throw new InvalidOperationException("The workbook does not contain a VBA project.");

                // Use reflection to obtain the DigitalSignature property (may not exist in older versions)
                PropertyInfo digitalSignatureProp = typeof(VbaProject).GetProperty("DigitalSignature", BindingFlags.Public | BindingFlags.Instance);
                if (digitalSignatureProp == null)
                    throw new NotSupportedException("DigitalSignature property is not available in the current Aspose.Cells version.");

                object digitalSignature = digitalSignatureProp.GetValue(vbaProject);
                if (digitalSignature == null)
                    throw new InvalidOperationException("The VBA project does not contain a digital certificate.");

                // Export the digital certificate to a MemoryStream via reflection
                MethodInfo exportMethod = digitalSignature.GetType().GetMethod("Export", new[] { typeof(Stream) });
                if (exportMethod == null)
                    throw new NotSupportedException("Export method is not available on DigitalSignature.");

                MemoryStream certificateStream = new MemoryStream();
                exportMethod.Invoke(digitalSignature, new object[] { certificateStream });
                certificateStream.Position = 0; // Reset for downstream processing

                return certificateStream;
            }
            catch (Exception ex)
            {
                // Wrap and rethrow to preserve stack trace
                throw new ApplicationException("Failed to export VBA digital certificate.", ex);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string excelPath = args.Length > 0 ? args[0] : "sample.xlsm";

                var exporter = new VbaCertificateExporter();
                using (MemoryStream certStream = exporter.ExportVbaCertificate(excelPath))
                {
                    // Example: save the certificate to a file
                    string outputPath = "certificate.cer";
                    using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        certStream.CopyTo(file);
                    }
                    Console.WriteLine($"Certificate exported successfully to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
