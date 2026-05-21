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
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (macro-enabled format will be used when saving)
                Workbook workbook = new Workbook();

                // Access the VBA project (it exists by default)
                VbaProject vbaProject = workbook.VbaProject;

                // Set optional project properties
                vbaProject.Name = "DemoVbaProject";
                vbaProject.Encoding = Encoding.UTF8;

                // Add a new class module to the VBA project
                int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
                VbaModule module = vbaProject.Modules[moduleIndex];

                // Insert VBA code into the module
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
                    certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);
                }
                catch (CryptographicException ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }

                // Create a DigitalSignature object
                DigitalSignature digitalSignature = new DigitalSignature(certificate, "VBA Project Signature", DateTime.Now);

                // Sign the VBA project
                try
                {
                    vbaProject.Sign(digitalSignature);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to sign VBA project: {ex.Message}");
                    return;
                }

                // Save the workbook as a macro-enabled file
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
                if (File.Exists(outputPath))
                {
                    Workbook verifyWorkbook = new Workbook(outputPath);
                    Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine($"Failed to locate saved workbook: {outputPath}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}