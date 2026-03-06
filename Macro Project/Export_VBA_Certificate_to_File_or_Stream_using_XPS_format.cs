using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Saving; // Namespace for XpsSaveOptions

class ExportVbaCertificateToXps
{
    static void Main()
    {
        // Path to the macro‑enabled workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, export its certificate raw data to a .cer file
        if (vbaProject != null && vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;
            if (certData != null && certData.Length > 0)
            {
                // Save the certificate to a file
                File.WriteAllBytes("VbaCertificate.cer", certData);
                Console.WriteLine("VBA certificate saved to VbaCertificate.cer");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project.");
        }

        // Create XPS save options (using the non‑obsolete constructor)
        XpsSaveOptions xpsOptions = new XpsSaveOptions
        {
            // Example option: put each sheet on a separate page
            OnePagePerSheet = true,
            // Set a default font to avoid missing‑font issues
            DefaultFont = "Arial"
        };

        // Save the workbook as XPS to a file
        string xpsFilePath = "WorkbookExport.xps";
        workbook.Save(xpsFilePath, xpsOptions);
        Console.WriteLine($"Workbook saved as XPS to {xpsFilePath}");

        // Additionally, demonstrate saving to a memory stream
        using (MemoryStream xpsStream = new MemoryStream())
        {
            // Save to stream using the same XpsSaveOptions
            workbook.Save(xpsStream, xpsOptions);

            // Optionally write the stream content to another file
            File.WriteAllBytes("WorkbookExportFromStream.xps", xpsStream.ToArray());
            Console.WriteLine("Workbook XPS also saved from memory stream.");
        }
    }
}