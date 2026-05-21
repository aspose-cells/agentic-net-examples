using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class CheckVbaSignature
{
    static void Main()
    {
        try
        {
            // Create a new macro‑enabled workbook
            Workbook workbook = new Workbook();

            // Add a VBA module to initialise the VBA project
            int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
            workbook.VbaProject.Modules[moduleIndex].Codes =
                "Sub Demo()\r\n    MsgBox \"Hello\"\r\nEnd Sub";

            // Path and password for the signing certificate
            string certPath = "mycert.pfx";
            string certPassword = "certPassword";

            // Load the certificate and sign the VBA project if the file exists
            if (File.Exists(certPath) && workbook.VbaProject != null)
            {
                try
                {
                    // Load certificate (obsolete warning suppressed, still functional)
                    X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);
                    DigitalSignature vbaSignature = new DigitalSignature(certificate, "VBA Signature", DateTime.Now);
                    workbook.VbaProject.Sign(vbaSignature);
                    Console.WriteLine("VBA project signed successfully.");
                }
                catch (CryptographicException ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    Console.WriteLine("Proceeding without signing the VBA project.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error during signing: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Certificate file not found or VBA project unavailable. Skipping VBA signing.");
            }

            // Save the workbook to a memory stream in macro‑enabled format
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, SaveFormat.Xlsm);
                stream.Position = 0; // Reset for reading

                // Reload the workbook to verify signature persistence
                Workbook reloadedWorkbook = new Workbook(stream);
                Console.WriteLine("VBA Project IsSigned: " + reloadedWorkbook.VbaProject.IsSigned);
                Console.WriteLine("VBA Project IsValidSigned: " + reloadedWorkbook.VbaProject.IsValidSigned);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}