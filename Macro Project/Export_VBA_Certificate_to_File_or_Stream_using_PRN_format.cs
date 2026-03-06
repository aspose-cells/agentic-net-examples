using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaCertificateExportDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Get the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export certificate to a file in PRN format
                    string certFilePath = "VbaCertificate.prn";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // Export certificate to a memory stream
                    using (MemoryStream certStream = new MemoryStream(certData))
                    {
                        // Write the stream to another file to verify
                        string streamFilePath = "VbaCertificateFromStream.prn";
                        using (FileStream fileStream = new FileStream(streamFilePath, FileMode.Create, FileAccess.Write))
                        {
                            certStream.CopyTo(fileStream);
                        }
                        Console.WriteLine($"Certificate written from stream to file: {streamFilePath}");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed. No certificate to export.");
            }

            // Save the workbook
            workbook.Save("VbaCertificateDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaCertificateExportDemo.Run();
        }
    }
}