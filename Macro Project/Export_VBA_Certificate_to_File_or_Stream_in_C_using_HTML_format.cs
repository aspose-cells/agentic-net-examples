using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Load a workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm"; // replace with actual path
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, extract the certificate raw data
        if (vbaProject.IsSigned)
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
            Console.WriteLine("VBA project is not signed.");
        }

        // Create HTML save options (embed images as Base64 for a self‑contained HTML)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;

        // Save the workbook as an HTML file on disk
        workbook.Save("Workbook.html", htmlOptions);
        Console.WriteLine("Workbook saved as HTML file.");

        // Save the same HTML output to a memory stream
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0; // reset for reading

            // Optionally write the stream content to another file for verification
            using (FileStream file = new FileStream("WorkbookFromStream.html", FileMode.Create, FileAccess.Write))
            {
                htmlStream.CopyTo(file);
            }
            Console.WriteLine("Workbook HTML saved to memory stream and written to file.");
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