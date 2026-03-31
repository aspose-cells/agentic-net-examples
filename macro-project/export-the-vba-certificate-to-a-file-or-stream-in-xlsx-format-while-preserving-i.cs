using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains a signed VBA project
            string sourcePath = "SignedWorkbook.xlsm";

            // Load the workbook (create rule is used internally by the constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export certificate to a file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // Export certificate to a memory stream (example of stream export)
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        certStream.Position = 0; // Reset position for further use

                        // For demonstration, write the stream content back to another file
                        string streamExportPath = "VbaCertificateFromStream.cer";
                        using (FileStream fileStream = new FileStream(streamExportPath, FileMode.Create, FileAccess.Write))
                        {
                            certStream.CopyTo(fileStream);
                        }
                        Console.WriteLine($"Certificate saved from stream to file: {streamExportPath}");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }

            // Save the workbook in XLSX format (preserving workbook content, macros are not retained in XLSX)
            string xlsxPath = "ExportedWorkbook.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as XLSX: {xlsxPath}");
        }
    }
}