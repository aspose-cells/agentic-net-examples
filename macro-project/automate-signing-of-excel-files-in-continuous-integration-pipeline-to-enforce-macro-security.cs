using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class SignExcelMacro
{
    static void Main()
    {
        // Path to the macro‑enabled workbook that needs to be signed
        string inputPath = "input.xlsm";

        // Path where the signed workbook will be saved
        string outputPath = "signed_output.xlsm";

        // Path to the code‑signing certificate (PFX) and its password
        string certificatePath = "mycert.pfx";
        string certificatePassword = "password";

        // Load the workbook that contains the VBA project
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project inside the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the workbook actually contains a VBA project
        if (vbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Load the signing certificate (must contain a private key)
        X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword);

        // Create a digital signature object using the certificate
        DigitalSignature digitalSignature = new DigitalSignature(certificate, "CI Signed Macro", DateTime.UtcNow);

        // Sign the VBA project with the digital signature
        vbaProject.Sign(digitalSignature);

        // Save the workbook; the VBA project is now signed
        workbook.Save(outputPath, SaveFormat.Xlsm);

        // Verify that the signature was applied correctly
        Workbook verificationWorkbook = new Workbook(outputPath);
        Console.WriteLine("VBA project signed: " + verificationWorkbook.VbaProject.IsSigned);
        Console.WriteLine("Signature valid: " + verificationWorkbook.VbaProject.IsValidSigned);
    }
}