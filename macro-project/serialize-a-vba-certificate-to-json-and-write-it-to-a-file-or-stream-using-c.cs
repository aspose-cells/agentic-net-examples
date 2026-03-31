using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateJson
{
    public class SerializeVbaCertificate
    {
        public static void Run()
        {
            // Load a workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed and obtain the certificate raw data
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                byte[] certData = vbaProject.CertRawData;

                // Convert the raw certificate bytes to a Base64 string for JSON representation
                string base64Cert = Convert.ToBase64String(certData);

                // Create an anonymous object to hold the certificate data
                var jsonObject = new
                {
                    CertificateBase64 = base64Cert,
                    CertificateLength = certData.Length,
                    IsSigned = vbaProject.IsSigned,
                    IsValidSigned = vbaProject.IsValidSigned
                };

                // Serialize the object to a formatted JSON string
                string json = JsonSerializer.Serialize(jsonObject, new JsonSerializerOptions { WriteIndented = true });

                // Write the JSON to a file
                string outputPath = "VbaCertificate.json";
                File.WriteAllText(outputPath, json);

                Console.WriteLine($"VBA certificate serialized to JSON and saved to '{outputPath}'.");
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            SerializeVbaCertificate.Run();
        }
    }
}