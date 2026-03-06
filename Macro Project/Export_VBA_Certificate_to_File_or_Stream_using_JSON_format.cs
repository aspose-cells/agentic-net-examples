using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            // Path to the macro‑enabled workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject == null || !vbaProject.IsSigned)
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

            // Convert the binary certificate to a Base64 string for JSON representation
            string certBase64 = Convert.ToBase64String(certData);

            // Write the certificate data into the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("CertificateBase64");
            sheet.Cells["A2"].PutValue(certBase64);

            // Configure JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export cell values as strings (the Base64 string will be preserved)
                ExportAsString = true,
                // Include header row in the JSON output
                HasHeaderRow = true,
                // Optional: make the JSON more readable
                Indent = "  "
            };

            // Example 1: Save the JSON to a file
            string jsonFilePath = "VbaCertificate.json";
            workbook.Save(jsonFilePath, jsonOptions);
            Console.WriteLine($"Certificate exported to JSON file: {jsonFilePath}");

            // Example 2: Save the JSON to a memory stream
            using (MemoryStream jsonStream = new MemoryStream())
            {
                workbook.Save(jsonStream, SaveFormat.Json);
                jsonStream.Position = 0;
                using (StreamReader reader = new StreamReader(jsonStream))
                {
                    string jsonContent = reader.ReadToEnd();
                    Console.WriteLine("Certificate exported to JSON stream:");
                    Console.WriteLine(jsonContent);
                }
            }
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