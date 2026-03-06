using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Load a macro-enabled workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Save the certificate raw data to a file
            string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, vbaProject.CertRawData);
            Console.WriteLine($"Certificate saved to file: {certFilePath}");

            // Also demonstrate exporting the certificate to a memory stream
            using (MemoryStream certStream = new MemoryStream(vbaProject.CertRawData))
            {
                // Example usage: display stream length
                Console.WriteLine($"Certificate stream length: {certStream.Length} bytes");
                // Reset position if further processing is required
                certStream.Position = 0;
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
        }

        // Save the workbook (preserving macros) as an XLSM file
        string outputPath = "OutputWorkbook.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
        Console.WriteLine($"Workbook saved as macro-enabled file: {outputPath}");
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportVbaCertificate.Run();
    }
}