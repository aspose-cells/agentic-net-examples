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
            // Load a macro-enabled workbook that contains a signed VBA project
            Workbook workbook = new Workbook("SignedWorkbook.xlsm");

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Save the certificate to a .cer file if data is available
                if (certData != null && certData.Length > 0)
                {
                    File.WriteAllBytes("VbaCertificate.cer", certData);
                    Console.WriteLine("VBA certificate saved to VbaCertificate.cer");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
            }

            // Export the workbook itself to MHTML (MHT) format
            workbook.Save("WorkbookExport.mht", SaveFormat.MHtml);
            Console.WriteLine("Workbook saved as MHTML to WorkbookExport.mht");

            // Optionally, export the MHTML content to a memory stream
            using (MemoryStream mhtStream = new MemoryStream())
            {
                workbook.Save(mhtStream, SaveFormat.MHtml);
                // Reset stream position for any further processing
                mhtStream.Position = 0;
                Console.WriteLine($"MHTML content written to memory stream (length: {mhtStream.Length} bytes).");
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