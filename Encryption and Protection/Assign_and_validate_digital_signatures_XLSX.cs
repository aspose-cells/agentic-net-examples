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
            // Paths (replace with actual locations)
            string certificatePath = "mycert.pfx";
            string certificatePassword = "password";
            string signedFilePath = "SignedWorkbook.xlsx";

            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Document requiring digital signature");

            // 2. Load the X.509 certificate (must contain a private key)
            X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword);

            // 3. Create a digital signature using the certificate
            DigitalSignature signature = new DigitalSignature(certificate, "Approved by QA", DateTime.UtcNow);

            // 4. Add the signature to a collection
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // 5. Apply the digital signature to the workbook
            workbook.SetDigitalSignature(signatures);

            // 6. Save the signed workbook
            workbook.Save(signedFilePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // Validation section: load the signed workbook and verify signatures
            // -----------------------------------------------------------------

            // 7. Load the signed workbook
            Workbook signedWorkbook = new Workbook(signedFilePath);

            // 8. Check if the workbook reports being digitally signed
            Console.WriteLine("Workbook.IsDigitallySigned = " + signedWorkbook.IsDigitallySigned);

            // 9. Retrieve the digital signature collection from the workbook
            DigitalSignatureCollection loadedSignatures = signedWorkbook.GetDigitalSignature();

            // 10. Iterate through each signature and display validation info
            if (loadedSignatures != null)
            {
                foreach (DigitalSignature ds in loadedSignatures)
                {
                    Console.WriteLine("Signature Comments : " + ds.Comments);
                    Console.WriteLine("Signature Time     : " + ds.SignTime);
                    Console.WriteLine("Signature IsValid  : " + ds.IsValid);
                }
            }
            else
            {
                Console.WriteLine("No digital signatures found in the workbook.");
            }
        }
    }
}