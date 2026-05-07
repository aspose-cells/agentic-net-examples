using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Export the certificate if the VBA project is signed
        if (vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;
            if (certData != null && certData.Length > 0)
            {
                // Save certificate raw data to a file
                string certPath = "VbaCertificate.cer";
                File.WriteAllBytes(certPath, certData);
                Console.WriteLine($"Certificate saved to '{certPath}'. Length: {certData.Length}");
            }
            else
            {
                Console.WriteLine("Certificate data is empty.");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed; no certificate to export.");
        }

        // Save the workbook as an XLTX template (preserving VBA project if possible)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xltx);
        string outputPath = "WorkbookTemplate.xltx";
        workbook.Save(outputPath, saveOptions);
        Console.WriteLine($"Workbook saved as XLTX template to '{outputPath}'.");
    }
}