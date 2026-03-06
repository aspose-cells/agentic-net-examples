using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Load a macro‑enabled workbook that contains a signed VBA project
        Workbook workbook = new Workbook("SignedWorkbook.xlsm");

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, retrieve its certificate raw data
        if (vbaProject != null && vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;
            if (certData != null && certData.Length > 0)
            {
                // Save the certificate to a .cer file
                File.WriteAllBytes("VbaCertificate.cer", certData);
                Console.WriteLine("Certificate saved to VbaCertificate.cer");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed; no certificate to export.");
        }

        // Export the workbook itself to MHTML (MHT) format and save to disk
        workbook.Save("Workbook.mht", SaveFormat.MHtml);
        Console.WriteLine("Workbook saved as MHTML to Workbook.mht");

        // Additionally, export the workbook to a memory stream in MHTML format
        using (MemoryStream mhtStream = new MemoryStream())
        {
            workbook.Save(mhtStream, SaveFormat.MHtml);
            mhtStream.Position = 0; // Reset stream position for further use

            // Example: write the stream content to another file
            using (FileStream file = new FileStream("WorkbookFromStream.mht", FileMode.Create, FileAccess.Write))
            {
                mhtStream.CopyTo(file);
            }
            Console.WriteLine("Workbook saved from stream to WorkbookFromStream.mht");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportVbaCertificate.Run();
    }
}