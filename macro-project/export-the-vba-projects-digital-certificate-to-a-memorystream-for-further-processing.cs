// Title: Export a VBA project's digital certificate to a MemoryStream using Aspose.Cells for .NET (C#)
// Common Searches: C# export VBA project digital certificate to a MemoryStream with Aspose.Cells | How to retrieve a VBA macro signature from an .xlsm file using Aspose.Cells | Using reflection to access VBA signature object in an Aspose.Cells workbook | Save VBA project's digital certificate as a .cer file after exporting to a stream in .NET | Aspose.Cells extract VBA digital signature when the Signature property is not directly exposed
// Tags: export VBA certificate Aspose.Cells C# | retrieve VBA digital signature from .xlsm workbook | memory stream VBA signature export .NET | reflection access VBA signature property Aspose.Cells | handle missing VBA project signature Aspose.Cells | save exported certificate to .cer file C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureExport
{
    // The example loads an .xlsm workbook with Aspose.Cells, uses reflection to obtain the VBA project's Signature object, invokes its Export method to write the digital certificate into a MemoryStream, resets the stream for further use, and optionally writes the bytes to a .cer file while gracefully handling absent signatures and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsm";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook that may contain a VBA project
                Workbook workbook = new Workbook(inputPath);

                // Prepare a memory stream to hold the exported certificate
                using (MemoryStream certificateStream = new MemoryStream())
                {
                    // Use reflection to access the VbaProject.Signature property (may not exist in older versions)
                    object vbaSignature = null;
                    PropertyInfo signatureProp = workbook.VbaProject?.GetType().GetProperty("Signature", BindingFlags.Public | BindingFlags.Instance);
                    if (signatureProp != null)
                    {
                        vbaSignature = signatureProp.GetValue(workbook.VbaProject);
                    }

                    if (vbaSignature != null)
                    {
                        // Use reflection to invoke the Export method on the signature object
                        MethodInfo exportMethod = vbaSignature.GetType().GetMethod("Export", new[] { typeof(Stream) });
                        if (exportMethod != null)
                        {
                            exportMethod.Invoke(vbaSignature, new object[] { certificateStream });

                            // Reset the stream position for subsequent reading
                            certificateStream.Position = 0;

                            // Example: save the certificate to a file (optional)
                            const string certPath = "vba_certificate.cer";
                            File.WriteAllBytes(certPath, certificateStream.ToArray());
                            Console.WriteLine($"Digital certificate exported successfully to '{certPath}'.");
                        }
                        else
                        {
                            Console.WriteLine("The VBA signature object does not support Export operation.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("No digital signature is attached to the VBA project or the API does not expose it.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors (e.g., loading issues, export failures)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
