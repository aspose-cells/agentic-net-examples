using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateJson
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project (provided by the VbaProject class)
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Retrieve the raw certificate data (byte array)
                byte[] certData = vbaProject.CertRawData;

                // Convert the certificate bytes to a Base64 string for JSON representation
                string base64Cert = Convert.ToBase64String(certData);

                // Build a simple JSON structure containing the certificate data
                string json = $"{{\"CertificateBase64\":\"{base64Cert}\",\"Length\":{certData.Length}}}";

                // ---------- Write JSON to a file ----------
                string jsonFilePath = "VbaCertificate.json";
                File.WriteAllText(jsonFilePath, json, Encoding.UTF8);
                Console.WriteLine($"Certificate JSON saved to file: {jsonFilePath}");

                // ---------- Write JSON to a memory stream ----------
                using (MemoryStream jsonStream = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(jsonStream, Encoding.UTF8, 1024, true))
                {
                    writer.Write(json);
                    writer.Flush();

                    // Reset stream position for further processing if needed
                    jsonStream.Position = 0;

                    // Example: read back the JSON from the stream
                    using (StreamReader reader = new StreamReader(jsonStream, Encoding.UTF8, true, 1024, true))
                    {
                        string jsonFromStream = reader.ReadToEnd();
                        Console.WriteLine("Certificate JSON read from stream:");
                        Console.WriteLine(jsonFromStream);
                    }
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
            }

            // Save the workbook if any modifications were made (uses the provided save rule)
            // In this example we don't modify the workbook, but the save call demonstrates the rule usage.
            // workbook.Save("ProcessedWorkbook.xlsm");
        }
    }
}