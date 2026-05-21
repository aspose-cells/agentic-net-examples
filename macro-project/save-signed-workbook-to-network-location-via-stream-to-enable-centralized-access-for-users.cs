using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

class SaveSignedWorkbookToNetwork
{
    static void Main()
    {
        try
        {
            // Verify certificate file exists
            string certPath = @"C:\Certificates\sample.pfx"; // adjust path as needed
            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Load certificate data
            string certPassword = "123456"; // certificate password
            byte[] certData = File.ReadAllBytes(certPath);
            DigitalSignature signature = new DigitalSignature(certData, certPassword, "Demo Signer", DateTime.UtcNow);
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);

            // Create workbook and add data
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Signed Workbook");
                sheet.Cells["B1"].PutValue(DateTime.Now);

                // Apply digital signature
                workbook.SetDigitalSignature(signatures);

                // Define UNC network path
                string networkFilePath = @"\\Server\SharedFolder\SignedWorkbook.xlsx";

                // Ensure target directory exists
                string networkDir = Path.GetDirectoryName(networkFilePath);
                if (!Directory.Exists(networkDir))
                {
                    Directory.CreateDirectory(networkDir);
                }

                // Save signed workbook to network location via stream
                using (FileStream networkStream = new FileStream(networkFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Save(networkStream, SaveFormat.Xlsx);
                }
            }

            Console.WriteLine("Signed workbook successfully saved to network location.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}