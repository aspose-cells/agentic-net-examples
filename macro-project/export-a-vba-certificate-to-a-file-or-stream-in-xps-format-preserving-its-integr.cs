using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Rendering;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificateToXps
    {
        public static void Run()
        {
            // Path to the macro‑enabled workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // If the VBA project is signed, extract the certificate raw data
            if (vbaProject != null && vbaProject.IsSigned)
            {
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Save the certificate to a .cer file (binary format)
                    File.WriteAllBytes("VbaCertificate.cer", certData);
                    Console.WriteLine($"Certificate saved (size: {certData.Length} bytes).");
                }
                else
                {
                    Console.WriteLine("Certificate raw data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed; no certificate to export.");
            }

            // Create XPS save options (preserve layout, one page per sheet)
            XpsSaveOptions xpsOptions = new XpsSaveOptions
            {
                OnePagePerSheet = true,
                DefaultFont = "Arial"
            };

            // Save the workbook as XPS using the options
            workbook.Save("SignedWorkbook.xps", xpsOptions);
            Console.WriteLine("Workbook saved as XPS.");

            // Optional: also save the XPS to a memory stream for further processing
            using (MemoryStream xpsStream = new MemoryStream())
            {
                workbook.Save(xpsStream, xpsOptions);
                // Reset stream position if needed
                xpsStream.Position = 0;
                // Example: write stream to a file (demonstrates stream usage)
                using (FileStream file = new FileStream("SignedWorkbook_Stream.xps", FileMode.Create, FileAccess.Write))
                {
                    xpsStream.CopyTo(file);
                }
                Console.WriteLine("Workbook XPS also saved via stream.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificateToXps.Run();
        }
    }
}