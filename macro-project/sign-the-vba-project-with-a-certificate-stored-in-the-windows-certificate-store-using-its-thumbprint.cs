using System;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class SignVbaProjectWithStoreCertificate
{
    static void Main()
    {
        // Paths for the input macro-enabled workbook and the signed output workbook
        string inputPath = "MacroWorkbook.xlsm";
        string outputPath = "MacroWorkbookSigned.xlsm";

        // Load the workbook that contains a VBA project
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        // Thumbprint of the certificate stored in the Windows certificate store
        // Replace the placeholder with the actual thumbprint (no spaces, case-insensitive)
        string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT".Replace(" ", "").ToUpperInvariant();

        // Open the current user's Personal (My) certificate store and locate the certificate by thumbprint
        X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        X509Certificate2 certificate = null;
        foreach (X509Certificate2 cert in store.Certificates)
        {
            if (!string.IsNullOrEmpty(cert.Thumbprint) &&
                cert.Thumbprint.Equals(thumbprint, StringComparison.OrdinalIgnoreCase))
            {
                certificate = cert;
                break;
            }
        }
        store.Close();

        if (certificate == null)
        {
            Console.WriteLine("Certificate with the specified thumbprint was not found in the store.");
            return;
        }

        // Create a digital signature using the located certificate
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,
            "Signed by certificate from Windows store",
            DateTime.Now);

        // Sign the VBA project
        vbaProject.Sign(digitalSignature);

        // Save the signed workbook as a macro-enabled file
        workbook.Save(outputPath, SaveFormat.Xlsm);

        // Optional verification after saving
        Workbook verifyWorkbook = new Workbook(outputPath);
        Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
        Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
    }
}