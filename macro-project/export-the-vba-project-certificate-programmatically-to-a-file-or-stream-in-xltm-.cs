using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Load a workbook that already contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm"; // replace with actual path
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the project is signed, retrieve the certificate raw data
        if (vbaProject != null && vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;
            if (certData != null && certData.Length > 0)
            {
                // Export the certificate to a .cer file
                File.WriteAllBytes("VbaCertificate.cer", certData);
                Console.WriteLine($"Certificate exported, length: {certData.Length} bytes.");
            }
            else
            {
                Console.WriteLine("Certificate data is empty.");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed or not present.");
        }

        // Save the workbook as a macro‑enabled template (XLTM)
        workbook.Save("ExportedTemplate.xltm", SaveFormat.Xltm);
        Console.WriteLine("Workbook saved as XLTM template.");
    }
}