// Title: C# – Create Macro‑Enabled XLSM Workbook, Add VBA Class Module, and Digitally Sign VBA Project with Aspose.Cells
// Description: Demonstrates how to generate a new XLSM workbook, insert a VBA class module, load an X.509 .pfx certificate, create a DigitalSignature, sign the VBA project, save the file, and verify the signature using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA signing | C# add VBA module | macro enabled workbook programmatically | digital signature XLSM | load .pfx certificate C# | verify VBA project signature | Aspose.Cells SaveFormat.Xlsm
// Common Searches: sign VBA project with Aspose.Cells C# | add class module to XLSM using .NET | digital signature for macro enabled workbook | load .pfx certificate in C# Aspose.Cells | check if VBA project is signed after save
// Developer Intent: Generate an XLSM file, embed a VBA class module, and apply a trusted digital signature to the VBA project programmatically.
// Use Cases: Automated creation of secure macro‑enabled reports for corporate distribution. | Ensuring VBA code is trusted to suppress macro security warnings in enterprise environments. | Validating the integrity of VBA projects before they enter a processing pipeline.
// AI Prompts: Write C# code that adds several VBA modules to a workbook and signs the project using a certificate stored in Azure Key Vault with Aspose.Cells. | Explain error‑handling strategies for loading .pfx certificates and falling back to a self‑signed certificate when signing VBA code. | Provide a step‑by‑step guide to verify a VBA project's digital signature after opening the workbook with Aspose.Cells.

using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Demonstrates how to generate a new XLSM workbook, insert a VBA class module, load an X.509 .pfx certificate, create a DigitalSignature, sign the VBA project, save the file, and verify the signature using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook (default format is XLSX)
                Workbook workbook = new Workbook();

                // Access the VBA project (it exists by default)
                VbaProject vbaProject = workbook.VbaProject;

                // Set VBA project properties (optional)
                vbaProject.Name = "DemoVbaProject";
                vbaProject.Encoding = Encoding.UTF8;

                // Add a new class module to the VBA project
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");

                // Retrieve the added module and set its VBA code
                VbaModule module = vbaProject.Modules[moduleIndex];
                module.Codes = @"Sub HelloWorld()
    MsgBox ""Hello from VBA!""
End Sub";

                // Load a digital certificate (replace with your actual .pfx path and password)
                string certPath = "MyCertificate.pfx";
                string certPassword = "password";

                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                X509Certificate2 certificate;
                try
                {
                    // Load certificate (obsolete warning is acceptable for demo purposes)
                    certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.DefaultKeySet);
                }
                catch (CryptographicException ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error loading certificate: {ex.Message}");
                    return;
                }

                // Create a digital signature object for the VBA project
                DigitalSignature vbaSignature = new DigitalSignature(certificate, "VBA Project Signature", DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(vbaSignature);

                // Save the workbook as a macro-enabled file (XLSM) to preserve VBA and its signature
                string outputPath = "SignedVbaWorkbook.xlsm";

                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsm);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                    return;
                }

                // Verify signing status
                Workbook verifyWorkbook = new Workbook(outputPath);
                Console.WriteLine("VBA Project Signed: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("Signature Valid: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
