using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

public static class ExportVbaCertificateToMhtml
{
    public static void Run()
    {
        // Load a macro‑enabled workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, extract the certificate raw data
        if (vbaProject != null && vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;
            if (certData != null && certData.Length > 0)
            {
                // Save the certificate to a .cer file
                File.WriteAllBytes("VbaCertificate.cer", certData);
            }
        }

        // Save the workbook as MHTML to a file
        workbook.Save("WorkbookExport.mht", SaveFormat.MHtml);

        // Also demonstrate saving to a memory stream (MHTML format)
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            workbook.Save(mhtmlStream, SaveFormat.MHtml);
            mhtmlStream.Position = 0; // Reset for further use

            // Example: write the stream content to another file
            using (FileStream file = new FileStream("WorkbookExportFromStream.mht", FileMode.Create, FileAccess.Write))
            {
                mhtmlStream.CopyTo(file);
            }
        }
    }
}

public class Program
{
    public static void Main()
    {
        ExportVbaCertificateToMhtml.Run();
    }
}