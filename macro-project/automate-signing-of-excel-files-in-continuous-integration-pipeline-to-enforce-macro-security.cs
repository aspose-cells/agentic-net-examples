using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class ExcelMacroSigner
{
    // Signs the VBA project of a macro‑enabled workbook using a digital certificate.
    public static void SignWorkbook(string inputPath, string outputPath, string certPath, string certPassword)
    {
        try
        {
            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException("Input workbook not found.", inputPath);

            // Verify that the certificate file exists.
            if (!File.Exists(certPath))
                throw new FileNotFoundException("Certificate file not found.", certPath);

            // Load the macro‑enabled workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project; throw if none exists.
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
                throw new InvalidOperationException("The workbook does not contain a VBA project.");

            // Load the signing certificate (must contain a private key).
            X509Certificate2 certificate = new X509Certificate2(certPath, certPassword);

            // Create a digital signature with comments and the current UTC time.
            DigitalSignature digitalSignature = new DigitalSignature(certificate, "CI Pipeline Signature", DateTime.UtcNow);

            // Apply the signature to the VBA project.
            vbaProject.Sign(digitalSignature);

            // Save the signed workbook in macro‑enabled format.
            workbook.Save(outputPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error signing workbook: {ex.Message}");
            throw; // Re‑throw to allow caller to handle if needed.
        }
    }

    static void Main()
    {
        // Paths for input workbook, output signed workbook, and the certificate.
        string inputWorkbook = "input.xlsm";
        string signedWorkbook = "signed_output.xlsm";
        string certificatePath = "ci_certificate.pfx";
        string certificatePassword = "yourPassword";

        try
        {
            // Perform signing.
            SignWorkbook(inputWorkbook, signedWorkbook, certificatePath, certificatePassword);

            // Verify that the signature was applied.
            Workbook verificationWorkbook = new Workbook(signedWorkbook);
            Console.WriteLine("VBA Project IsSigned: " + verificationWorkbook.VbaProject.IsSigned);
            Console.WriteLine("VBA Project IsValidSigned: " + verificationWorkbook.VbaProject.IsValidSigned);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}