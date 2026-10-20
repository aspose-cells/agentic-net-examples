// Title: Export a VBA project's digital certificate to a Base64 string using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export VBA project certificate as Base64 string | How to read VBA macro digital signature from an Excel file with Aspose.Cells | Retrieve VBA project signature certificate bytes in .NET | Convert VBA project certificate to Base64 using reflection in C# | Aspose.Cells get VBA project certificate for configuration
// Tags: Aspose.Cells export VBA certificate to Base64 | C# retrieve VBA project signature certificate | reflection access VBA signature property Aspose.Cells | Excel workbook VBA certificate extraction .NET | convert certificate byte array to Base64 C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject

namespace ExportVbaCertificateApp
{
    // The example loads an Excel workbook, accesses its embedded VBA project, uses reflection to obtain the Signature object's Certificate byte array, converts the certificate to a Base64 string, and outputs the result.
    class ExportVbaCertificate
    {
        static void Main()
        {
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the Excel workbook
                Workbook workbook = new Workbook(filePath);

                // Access the VBA project embedded in the workbook
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject == null)
                {
                    Console.WriteLine("No VBA project found in the workbook.");
                    return;
                }

                // Use reflection to obtain the Signature property (may not exist in older versions)
                PropertyInfo signatureProp = vbaProject.GetType().GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
                if (signatureProp == null)
                {
                    Console.WriteLine("VBA signature information is not available in this Aspose.Cells version.");
                    return;
                }

                object signatureObj = signatureProp.GetValue(vbaProject);
                if (signatureObj == null)
                {
                    Console.WriteLine("No VBA signature found in the workbook.");
                    return;
                }

                // Use reflection to obtain the Certificate byte[] from the signature object
                PropertyInfo certificateProp = signatureObj.GetType().GetProperty("Certificate", BindingFlags.Public | BindingFlags.Instance);
                if (certificateProp == null)
                {
                    Console.WriteLine("Certificate property is not available in the VBA signature.");
                    return;
                }

                byte[] certificateBytes = certificateProp.GetValue(signatureObj) as byte[];
                if (certificateBytes != null && certificateBytes.Length > 0)
                {
                    // Convert the certificate byte array to a Base64 string
                    string base64Certificate = Convert.ToBase64String(certificateBytes);
                    Console.WriteLine("VBA Certificate (Base64):");
                    Console.WriteLine(base64Certificate);
                }
                else
                {
                    Console.WriteLine("No VBA certificate found in the workbook.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
