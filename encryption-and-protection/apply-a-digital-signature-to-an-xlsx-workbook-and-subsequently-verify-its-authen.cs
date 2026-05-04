using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsDigitalSignatureDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the PFX certificate file and its password
            string certificatePath = "mycert.pfx";
            string certificatePassword = "password";

            // Path where the signed workbook will be saved
            string signedWorkbookPath = "SignedWorkbook.xlsx";

            // -------------------------------------------------
            // Create a new workbook and add some sample data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Document to be signed");

            // -------------------------------------------------
            // Load the X.509 certificate (must contain a private key)
            // -------------------------------------------------
            X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword);

            // -------------------------------------------------
            // Create a digital signature using the certificate
            // -------------------------------------------------
            DigitalSignature signature = new DigitalSignature(
                certificate,               // certificate with private key
                "Demo signature",          // comments / purpose
                DateTime.UtcNow);          // sign time (UTC)

            // -------------------------------------------------
            // Put the signature into a collection (required by API)
            // -------------------------------------------------
            DigitalSignatureCollection signatureCollection = new DigitalSignatureCollection();
            signatureCollection.Add(signature);

            // -------------------------------------------------
            // Apply the digital signature to the workbook
            // -------------------------------------------------
            workbook.AddDigitalSignature(signatureCollection);

            // -------------------------------------------------
            // Save the signed workbook to disk
            // -------------------------------------------------
            workbook.Save(signedWorkbookPath);

            // -------------------------------------------------
            // Load the signed workbook for verification
            // -------------------------------------------------
            Workbook signedWorkbook = new Workbook(signedWorkbookPath);

            // -------------------------------------------------
            // Verify that the workbook reports being digitally signed
            // -------------------------------------------------
            Console.WriteLine($"Is workbook digitally signed? {signedWorkbook.IsDigitallySigned}");

            // -------------------------------------------------
            // Retrieve the digital signature collection from the file
            // -------------------------------------------------
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            // -------------------------------------------------
            // Iterate through each signature and check its validity
            // -------------------------------------------------
            foreach (DigitalSignature ds in loadedSignatures)
            {
                Console.WriteLine($"Comments   : {ds.Comments}");
                Console.WriteLine($"Sign Time  : {ds.SignTime:u}");
                Console.WriteLine($"Is Valid   : {ds.IsValid}");
                Console.WriteLine(new string('-', 30));
            }
        }
    }
}