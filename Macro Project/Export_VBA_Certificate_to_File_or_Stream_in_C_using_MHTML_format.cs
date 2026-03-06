using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificateMhtml
{
    static void Main()
    {
        // Path to the macro‑enabled workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, retrieve its certificate raw data
        if (vbaProject != null && vbaProject.IsSigned)
        {
            byte[] certData = vbaProject.CertRawData;

            // Save the certificate data to a .cer file if it exists
            if (certData != null && certData.Length > 0)
            {
                File.WriteAllBytes("VbaCertificate.cer", certData);
            }
        }

        // Export the workbook to MHTML (MHT) format and save to a file
        workbook.Save("WorkbookExport.mht", SaveFormat.MHtml);

        // Additionally, export the workbook to MHTML format into a memory stream
        using (MemoryStream mhtmlStream = new MemoryStream())
        {
            workbook.Save(mhtmlStream, SaveFormat.MHtml);

            // For demonstration, write the stream content to another file
            File.WriteAllBytes("WorkbookExportFromStream.mht", mhtmlStream.ToArray());
        }
    }
}