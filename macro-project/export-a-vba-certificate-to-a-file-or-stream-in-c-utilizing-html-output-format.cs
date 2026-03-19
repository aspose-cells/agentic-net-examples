using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Saving;

public class ExportVbaCertificateDemo
{
    public static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Export the VBA certificate if the project is signed
        if (vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;

            if (certData != null && certData.Length > 0)
            {
                // Save the raw certificate data to a .cer file
                File.WriteAllBytes("VbaCertificate.cer", certData);
                Console.WriteLine("Certificate saved to VbaCertificate.cer");
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

        // Create HTML save options (default settings)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as an HTML file on disk
        workbook.Save("Workbook.html", htmlOptions);
        Console.WriteLine("Workbook saved as HTML to Workbook.html");

        // Additionally, save the HTML output to a memory stream
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);

            // Optionally write the stream content to another file for verification
            File.WriteAllBytes("WorkbookFromStream.html", htmlStream.ToArray());
            Console.WriteLine("Workbook HTML saved from stream to WorkbookFromStream.html");
        }
    }
}