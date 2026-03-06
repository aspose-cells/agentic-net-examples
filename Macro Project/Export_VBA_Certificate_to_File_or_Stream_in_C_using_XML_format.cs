using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            // Load a workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Create Workbook instance (load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is not signed. No certificate to export.");
                return;
            }

            // Retrieve the raw certificate data (byte array)
            byte[] certData = vbaProject.CertRawData;

            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // Export certificate to a physical file (binary .cer)
            string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, certData);
            Console.WriteLine($"Certificate saved to file: {certFilePath}");

            // Export certificate to a MemoryStream (XML format wrapper)
            using (MemoryStream xmlStream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                writer.WriteLine("<VbaCertificate>");
                string base64Cert = Convert.ToBase64String(certData);
                writer.WriteLine($"  <RawData>{base64Cert}</RawData>");
                writer.WriteLine("</VbaCertificate>");
                writer.Flush();

                xmlStream.Position = 0;

                string xmlFilePath = "VbaCertificate.xml";
                using (FileStream fileStream = new FileStream(xmlFilePath, FileMode.Create, FileAccess.Write))
                {
                    xmlStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Certificate exported as XML to: {xmlFilePath}");
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