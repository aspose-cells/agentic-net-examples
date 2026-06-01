using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigning
{
    public class SignVbaProjectWithStoreCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a VBA project (must be macro‑enabled)
                string workbookPath = "InputWorkbook.xlsm";

                // Verify the input file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input workbook not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Ensure the workbook has a VBA project; if not, create one by saving as .xlsm and reloading
                if (workbook.VbaProject == null)
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsm");
                    workbook.Save(tempPath, SaveFormat.Xlsm);
                    workbook = new Workbook(tempPath);
                    File.Delete(tempPath);
                }

                // Retrieve the certificate from the Windows certificate store by thumbprint
                string thumbprint = "YOUR_CERTIFICATE_THUMBPRINT".Replace(" ", "").ToUpperInvariant(); // replace with actual thumbprint
                X509Certificate2 certificate = null;

                using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
                {
                    store.Open(OpenFlags.ReadOnly);
                    var certCollection = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                    if (certCollection.Count > 0)
                    {
                        certificate = certCollection[0];
                    }
                }

                if (certificate == null)
                {
                    Console.WriteLine("Certificate with the specified thumbprint was not found in the store.");
                    return;
                }

                // Create a digital signature using the retrieved certificate
                DigitalSignature digitalSignature = new DigitalSignature(
                    certificate,
                    "VBA Project Signature",
                    DateTime.Now);

                // Sign the VBA project
                VbaProject vbaProject = workbook.VbaProject;
                vbaProject.Sign(digitalSignature);

                // Save the signed workbook
                string signedWorkbookPath = "SignedWorkbook.xlsm";
                workbook.Save(signedWorkbookPath, SaveFormat.Xlsm);

                // Optional: verify the signature
                Workbook verifyWorkbook = new Workbook(signedWorkbookPath);
                Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
                Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SignVbaProjectWithStoreCertificate.Run();
        }
    }
}