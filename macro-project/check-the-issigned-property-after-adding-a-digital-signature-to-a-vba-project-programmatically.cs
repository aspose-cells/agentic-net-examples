// Title: Check VBA Project IsSigned After Adding a Digital Signature with Aspose.Cells for .NET (C#)
// Description: Creates a temporary macro‑enabled workbook, loads an X509 certificate, signs the workbook's VBA project using Aspose.Cells.DigitalSignatures, prints the IsSigned and IsValidSigned flags, saves the file to a memory stream, reloads it, and confirms that the signature status persists after the round‑trip.
// Keywords: Aspose.Cells VBA digital signature | C# check IsSigned property | verify VBA signature persistence | sign macro-enabled workbook programmatically | IsValidSigned Aspose.Cells | load X509 certificate C# | VBA project signing example | Aspose.Cells .NET digital signature
// Common Searches: How to check if a VBA project is signed with Aspose.Cells .NET | Verify VBA digital signature after saving workbook in C# | Programmatically sign a macro‑enabled workbook and read IsSigned | Persist VBA signature when saving to a memory stream using Aspose.Cells | Aspose.Cells C# example for VBA project signing and validation
// Developer Intent: The developer needs to sign a VBA project programmatically and ensure that the IsSigned flag remains true after the workbook is saved and reloaded.
// Use Cases: Sign a newly created macro‑enabled workbook and validate the signature before distribution. | Load a workbook from a stream, check IsSigned and IsValidSigned to detect tampering. | Automate logging of signature validation results in a batch processing pipeline.
// AI Prompts: Generate C# code that uses Aspose.Cells to sign a VBA project, then reload the workbook and verify IsSigned and IsValidSigned properties. | Explain how to handle missing or invalid X509 certificate files when signing a VBA project with Aspose.Cells. | Show how to remove a digital signature from a VBA project and confirm that IsSigned becomes false using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;
using System.Security.Cryptography.X509Certificates;

// Creates a temporary macro‑enabled workbook, loads an X509 certificate, signs the workbook's VBA project using Aspose.Cells.DigitalSignatures, prints the IsSigned and IsValidSigned flags, saves the file to a memory stream, reloads it, and confirms that the signature status persists after the round‑trip.
class CheckVbaProjectSignature
{
    static void Main()
    {
        const string tempFile = "temp.xlsm";
        const string certificatePath = "MyCertificate.pfx";
        const string certificatePassword = "password";

        try
        {
            // Create a temporary macro‑enabled workbook
            Workbook tempWorkbook = new Workbook();
            tempWorkbook.Save(tempFile, SaveFormat.Xlsm);

            // Load the workbook (now it contains a VBA project placeholder)
            Workbook workbook = new Workbook(tempFile);

            // Load the certificate if the file exists
            X509Certificate2 certificate = null;
            if (File.Exists(certificatePath))
            {
                try
                {
                    certificate = new X509Certificate2(certificatePath, certificatePassword);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Certificate file '{certificatePath}' not found. Skipping signing.");
            }

            // Sign the VBA project if both certificate and VBA project are available
            VbaProject vbaProject = workbook.VbaProject;
            if (certificate != null && vbaProject != null)
            {
                DigitalSignature vbaSignature = new DigitalSignature(certificate, "VBA Project Signature", DateTime.Now);
                vbaProject.Sign(vbaSignature);
                Console.WriteLine("After signing - IsSigned: " + vbaProject.IsSigned);
                Console.WriteLine("After signing - IsValidSigned: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project not available or certificate not loaded. Skipping signing.");
            }

            // Verify that the signature persists after saving to a stream
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0; // Reset stream for reading

                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;
                if (reloadedVba != null)
                {
                    Console.WriteLine("After reload - IsSigned: " + reloadedVba.IsSigned);
                    Console.WriteLine("After reload - IsValidSigned: " + reloadedVba.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("Reloaded workbook does not contain a VBA project.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file
            if (File.Exists(tempFile))
            {
                try
                {
                    File.Delete(tempFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }
    }
}
