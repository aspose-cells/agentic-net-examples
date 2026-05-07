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
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Load the workbook (create rule: Workbook constructor with file path)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project (property rule)
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data (property rule)
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate to a PRN file, preserving its binary integrity
                    string prnFilePath = "VbaCertificate.prn";
                    File.WriteAllBytes(prnFilePath, certData);
                    Console.WriteLine($"Certificate exported to PRN file: {prnFilePath}");

                    // Additionally, demonstrate exporting to a memory stream and then to a PRN file
                    using (MemoryStream ms = new MemoryStream())
                    {
                        // Write the certificate bytes into the memory stream
                        ms.Write(certData, 0, certData.Length);
                        ms.Position = 0; // Reset position for reading

                        // Save the stream content to another PRN file
                        string prnFromStreamPath = "VbaCertificateFromStream.prn";
                        using (FileStream fileStream = new FileStream(prnFromStreamPath, FileMode.Create, FileAccess.Write))
                        {
                            ms.CopyTo(fileStream);
                        }

                        Console.WriteLine($"Certificate exported from stream to PRN file: {prnFromStreamPath}");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed. No certificate to export.");
            }
        }
    }
}