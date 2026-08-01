// Title: C# – Sign a VBA Project in an .xlsm Workbook with a PFX Certificate using Aspose.Cells
// Description: Load a macro‑enabled workbook, create an X509Certificate2 from a password‑protected PFX file, build a DigitalSignature, sign the VBA project, save the workbook, and optionally verify the signature status—all with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA signing | C# load X509Certificate2 PFX | digital signature Excel macro | sign .xlsm VBA project | verify VBA signature Aspose | programmatic code signing Excel | certificate based macro protection
// Common Searches: how to sign a VBA project in an xlsm file using Aspose.Cells | C# load PFX certificate and sign Excel macro workbook | verify VBA project signature after saving with Aspose.Cells | programmatically apply digital signature to Excel macros | Aspose.Cells example for signing VBA with X509Certificate2
// Developer Intent: Create an X509Certificate2 from a PFX file and use it to digitally sign the VBA project of a macro‑enabled workbook.
// Use Cases: Apply a corporate code‑signing certificate to protect VBA macros before distribution. | Automate signing of generated macro‑enabled reports in CI/CD pipelines to guarantee macro integrity. | Validate that a signed VBA project remains intact after the workbook is saved and reopened.
// AI Prompts: Generate C# code that loads a password‑protected PFX certificate and signs a VBA project with Aspose.Cells, including robust error handling. | Explain step‑by‑step how to verify a VBA project's digital signature after saving the workbook using Aspose.Cells. | Recommend a secure method for storing and retrieving the PFX password in a .NET application while signing VBA projects.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    // Load a macro‑enabled workbook, create an X509Certificate2 from a password‑protected PFX file, build a DigitalSignature, sign the VBA project, save the workbook, and optionally verify the signature status—all with Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path to the macro-enabled workbook that contains a VBA project
                string workbookPath = "SampleWithVba.xlsm";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject != null)
                {
                    // Path to the certificate file
                    string pfxPath = "MyCertificate.pfx";

                    // Verify that the certificate file exists
                    if (!File.Exists(pfxPath))
                    {
                        Console.WriteLine($"Certificate file not found: {pfxPath}");
                        return;
                    }

                    // Load the certificate from a PFX file with its password
                    string pfxPassword = "certPassword";
                    X509Certificate2 certificate = new X509Certificate2(pfxPath, pfxPassword);

                    // Create a digital signature using the certificate
                    DigitalSignature digitalSignature = new DigitalSignature(
                        certificate,
                        "Signed by Aspose.Cells VBA demo",
                        DateTime.Now);

                    // Sign the VBA project
                    vbaProject.Sign(digitalSignature);

                    // Save the signed workbook as a macro-enabled file
                    string signedWorkbookPath = "SignedWorkbook.xlsm";
                    workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);

                    // Optional: Verify the signature after saving
                    Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                    Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                    Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a VBA project.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
