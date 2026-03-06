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
            // Load a workbook that contains a signed VBA project.
            // Replace the path with the actual file you want to process.
            string workbookPath = "SignedWorkbook.xlsm";
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Create (load) the workbook.
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed.
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate data available.");
                return;
            }

            // Retrieve the raw certificate data (binary).
            byte[] certData = vbaProject.CertRawData;

            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // Export to a text file (binary data saved with .txt extension)
            string txtFilePath = "VbaCertificate.txt";
            File.WriteAllBytes(txtFilePath, certData);
            Console.WriteLine($"Certificate saved to file: {txtFilePath}");

            // Export to a memory stream and then write the stream to another file
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(certData, 0, certData.Length);
                certStream.Position = 0;

                string streamFilePath = "VbaCertificateFromStream.txt";
                using (FileStream fileStream = new FileStream(streamFilePath, FileMode.Create, FileAccess.Write))
                {
                    certStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Certificate saved from stream to file: {streamFilePath}");
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