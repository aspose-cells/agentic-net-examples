// Title: Export a signed VBA project's X509 certificate to a .cer file and display its thumbprint using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to extract the signing X509Certificate2 from a VBA project and write it to a .cer file. | Retrieve the certificate of a signed VBA macro via reflection and output its thumbprint with Aspose.Cells. | Save the VBA project's signing certificate as a .cer file and print the thumbprint using the Aspose.Cells API.
// Common Searches: how to export VBA project signing certificate to .cer with Aspose.Cells | c# get thumbprint of signed Excel VBA macro certificate | extract X509Certificate2 from signed VBA project using reflection | save VBA macro certificate as file Aspose.Cells .NET | retrieve and display VBA project certificate thumbprint programmatically
// Tags: export vba signing certificate as cer Aspose.Cells | retrieve vba project certificate via reflection C# | display vba macro certificate thumbprint .NET | extract signed vba project certificate Aspose.Cells | save X509Certificate2 from excel vba project

using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an Excel workbook, accesses its VBA project, uses reflection to obtain the X509Certificate2 when the project is signed, exports the certificate to a .cer file, and prints the certificate's thumbprint.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Attempt to retrieve the certificate via reflection (covers versions without direct Signature property)
            X509Certificate2 cert = null;
            if (vbaProject != null && vbaProject.IsSigned)
            {
                try
                {
                    PropertyInfo signatureProp = vbaProject.GetType().GetProperty("Signature");
                    if (signatureProp != null)
                    {
                        object signatureObj = signatureProp.GetValue(vbaProject);
                        if (signatureObj != null)
                        {
                            PropertyInfo certProp = signatureObj.GetType().GetProperty("Certificate");
                            if (certProp != null)
                            {
                                cert = certProp.GetValue(signatureObj) as X509Certificate2;
                            }
                        }
                    }
                }
                catch (Exception reflEx)
                {
                    Console.WriteLine($"Reflection error while accessing signature: {reflEx.Message}");
                }
            }

            if (cert != null)
            {
                // Export the certificate to a .cer file
                byte[] certBytes = cert.Export(X509ContentType.Cert);
                string outputPath = "project.cer";

                try
                {
                    File.WriteAllBytes(outputPath, certBytes);
                    Console.WriteLine($"Certificate exported to '{outputPath}'.");
                    Console.WriteLine($"Certificate thumbprint: {cert.Thumbprint}");
                }
                catch (Exception ioEx)
                {
                    Console.WriteLine($"Failed to write certificate file: {ioEx.Message}");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project with an accessible certificate.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
