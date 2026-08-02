using System;
using System.IO;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsVbaSigningDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (default format is XLSX)
            Workbook workbook = new Workbook();

            // Access the VBA project (it exists by default)
            VbaProject vbaProject = workbook.VbaProject;

            // Optional: set project name and encoding
            vbaProject.Name = "MyVbaProject";
            vbaProject.Encoding = Encoding.UTF8;

            // Add a class module to the VBA project
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyModule");

            // Set VBA code for the newly added module
            VbaModule module = vbaProject.Modules[moduleIndex];
            module.Codes = "Sub Hello()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

            // Load a digital certificate (replace with your actual .pfx path and password)
            X509Certificate2 certificate = new X509Certificate2("MyCertificate.pfx", "certPassword");

            // Create a DigitalSignature object using the certificate
            DigitalSignature digitalSignature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);

            // Sign the VBA project with the digital signature
            vbaProject.Sign(digitalSignature);

            // Save the workbook as a macro-enabled file to preserve the VBA project
            string outputPath = "SignedVbaWorkbook.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);

            // Verify the signature by reloading the workbook
            Workbook verifyWorkbook = new Workbook(outputPath);
            Console.WriteLine("VBA Project IsSigned: " + verifyWorkbook.VbaProject.IsSigned);
            Console.WriteLine("VBA Project IsValidSigned: " + verifyWorkbook.VbaProject.IsValidSigned);
        }
    }
}