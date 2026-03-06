using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Saving;

public class ExportVbaCertificateDemo
{
    public static void Run()
    {
        // Path to a macro‑enabled workbook that contains a signed VBA project
        string sourcePath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and that certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Export certificate raw data to a physical file
            string certFilePath = "VbaCertificate.cer";
            File.WriteAllBytes(certFilePath, vbaProject.CertRawData);
            Console.WriteLine($"Certificate saved to file: {certFilePath} (Length: {vbaProject.CertRawData.Length} bytes)");

            // Export certificate raw data to a memory stream (example)
            using (MemoryStream certStream = new MemoryStream())
            {
                certStream.Write(vbaProject.CertRawData, 0, vbaProject.CertRawData.Length);
                certStream.Position = 0; // Reset for potential further reading
                Console.WriteLine($"Certificate written to memory stream (Size: {certStream.Length} bytes)");
                // Additional processing of the stream can be performed here
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or the certificate data is unavailable.");
        }

        // Save the workbook as OXPS (XPS) using XpsSaveOptions
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            OnePagePerSheet = true
        };

        string oxpsOutputPath = "WorkbookOutput.xps";
        workbook.Save(oxpsOutputPath, xpsOptions);
        Console.WriteLine($"Workbook saved as OXPS to: {oxpsOutputPath}");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        ExportVbaCertificateDemo.Run();
    }
}