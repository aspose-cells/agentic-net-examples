using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            // Load a workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"File not found: {signedWorkbookPath}");
                return;
            }

            // Create (load) the workbook using the provided load rule
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate data available.");
                return;
            }

            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;
            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // -------------------------------------------------
            // Export certificate to a binary .cer file (file)
            // -------------------------------------------------
            string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, certData);
            Console.WriteLine($"Certificate saved to file: {certFilePath}");

            // -------------------------------------------------
            // Export certificate to a memory stream (stream)
            // -------------------------------------------------
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(certData, 0, certData.Length);
                certStream.Position = 0; // Reset for reading if needed

                // Example: write the stream content to another file to verify
                string streamOutputPath = "VbaCertificateFromStream.cer";
                using (FileStream fileStream = new FileStream(streamOutputPath, FileMode.Create, FileAccess.Write))
                {
                    certStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Certificate written from stream to file: {streamOutputPath}");
            }

            // -------------------------------------------------
            // Export certificate data as JSON (base64 encoded)
            // -------------------------------------------------
            var jsonObject = new
            {
                CertificateBase64 = Convert.ToBase64String(certData),
                Length = certData.Length
            };

            string json = JsonSerializer.Serialize(jsonObject, new JsonSerializerOptions { WriteIndented = true });
            string jsonFilePath = "VbaCertificate.json";
            File.WriteAllText(jsonFilePath, json);
            Console.WriteLine($"Certificate exported to JSON file: {jsonFilePath}");

            // -------------------------------------------------
            // Optional: Save the workbook (demonstrating the save rule)
            // -------------------------------------------------
            string savedWorkbookPath = "WorkbookAfterExport.xlsm";
            workbook.Save(savedWorkbookPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {savedWorkbookPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}