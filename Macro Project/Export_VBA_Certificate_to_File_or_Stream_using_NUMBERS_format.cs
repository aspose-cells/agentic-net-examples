using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Load an existing workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is signed and certificate data is available
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Save the certificate raw data to a .cer file
            string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, vbaProject.CertRawData);
            Console.WriteLine($"Certificate saved to file: {certFilePath}");

            // Optionally, write the certificate data to a memory stream
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(vbaProject.CertRawData, 0, vbaProject.CertRawData.Length);
                certStream.Position = 0; // Reset for further reading if needed
                Console.WriteLine($"Certificate written to memory stream (length: {certStream.Length} bytes).");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
        }

        // Save the workbook in Numbers format
        string numbersFilePath = "WorkbookNumbers.numbers";
        workbook.Save(numbersFilePath, SaveFormat.Numbers);
        Console.WriteLine($"Workbook saved in Numbers format: {numbersFilePath}");
    }
}