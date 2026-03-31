using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateToNumbersDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // If certificate data exists, save it to a .cer file
                if (certData != null && certData.Length > 0)
                {
                    string certPath = "VbaCertificate.cer";
                    File.WriteAllBytes(certPath, certData);
                    Console.WriteLine($"VBA certificate saved to '{certPath}'.");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist; no certificate to export.");
            }

            // Save the workbook in Numbers format
            string numbersPath = "WorkbookNumbers.numbers";
            workbook.Save(numbersPath, SaveFormat.Numbers);
            Console.WriteLine($"Workbook saved in Numbers format to '{numbersPath}'.");

            // Optional: also demonstrate saving to a memory stream in Numbers format
            using (MemoryStream numbersStream = new MemoryStream())
            {
                workbook.Save(numbersStream, SaveFormat.Numbers);
                numbersStream.Position = 0;

                string streamOutputPath = "WorkbookNumbersFromStream.numbers";
                using (FileStream fileStream = new FileStream(streamOutputPath, FileMode.Create, FileAccess.Write))
                {
                    numbersStream.CopyTo(fileStream);
                }
                Console.WriteLine($"Workbook saved from stream to '{streamOutputPath}'.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificateToNumbersDemo.Run();
        }
    }
}