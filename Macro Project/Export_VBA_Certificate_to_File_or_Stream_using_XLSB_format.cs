using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main()
        {
            // Path to an existing workbook that contains a signed VBA project (macro‑enabled file)
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // If the VBA project is signed, export its certificate raw data to a .cer file
            if (vbaProject.IsSigned)
            {
                byte[] certData = vbaProject.CertRawData;
                if (certData != null && certData.Length > 0)
                {
                    File.WriteAllBytes("VbaCertificate.cer", certData);
                }
            }

            // Prepare XLSB save options
            XlsbSaveOptions xlsbOptions = new XlsbSaveOptions();

            // Save the workbook as an XLSB file using the options (file output)
            workbook.Save("ExportedWorkbook.xlsb", xlsbOptions);

            // Additionally, save the workbook to a memory stream in XLSB format
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.Save(stream, xlsbOptions);

                // Example: write the stream content to another file for verification
                File.WriteAllBytes("ExportedWorkbookFromStream.xlsb", stream.ToArray());
            }

            // Clean up
            workbook.Dispose();
        }
    }
}