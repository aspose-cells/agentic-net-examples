using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class ReSignVbaProjectDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string workbookPath = "OriginalWithVba.xlsm";
                const string certPath = "MyCertificate.pfx";
                const string certPassword = "certificatePassword";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Verify that the certificate file exists
                if (!File.Exists(certPath))
                {
                    Console.WriteLine($"Certificate file not found: {certPath}");
                    return;
                }

                // Load the workbook that contains a VBA project
                Workbook workbook = new Workbook(workbookPath);

                // Get the VBA project from the workbook
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject == null)
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                    return;
                }

                // Modify VBA code if there is at least one module
                if (vbaProject.Modules.Count > 0)
                {
                    VbaModule module = vbaProject.Modules[0];
                    // Append a comment to indicate that the code was changed
                    module.Codes += "\n' Modified by ReSignVbaProjectDemo";
                }

                // Load the signing certificate
                X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

                // Create a digital signature object
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "Re-signed after code changes",
                    DateTime.Now);

                // Sign the VBA project with the new digital signature
                vbaProject.Sign(digitalSignature);

                // Save the workbook (macro-enabled) to a memory stream
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.Save(stream, SaveFormat.Xlsm);
                    stream.Position = 0; // Reset stream position for reading

                    // Reload the workbook from the stream to verify signature status
                    Workbook verifyWorkbook = new Workbook(stream);
                    Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}