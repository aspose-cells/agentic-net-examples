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
            // Path to the source workbook that contains a signed VBA project
            string sourceWorkbookPath = "SignedWorkbook.xlsm";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(sourceWorkbookPath);

            // Access the VBA project (property rule)
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed
            bool isSigned = vbaProject.IsSigned;

            // Prepare a Base64 representation of the certificate raw data (if available)
            string certificateBase64 = string.Empty;
            if (isSigned)
            {
                byte[] certData = vbaProject.CertRawData;
                if (certData != null && certData.Length > 0)
                {
                    certificateBase64 = Convert.ToBase64String(certData);
                }
            }

            // Write required fields into the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("IsSigned");
            sheet.Cells["B1"].PutValue(isSigned);
            sheet.Cells["A2"].PutValue("CertificateBase64");
            sheet.Cells["B2"].PutValue(certificateBase64);

            // Export the worksheet to a CSV file (save rule)
            string csvFilePath = "VbaCertificateExport.csv";
            workbook.Save(csvFilePath, SaveFormat.Csv);

            // Additionally, export the same CSV content to a memory stream
            using (MemoryStream csvStream = new MemoryStream())
            {
                workbook.Save(csvStream, SaveFormat.Csv);
                // Reset stream position for any further processing
                csvStream.Position = 0;

                // Example: write the stream content to another file
                using (FileStream file = new FileStream("VbaCertificateExportFromStream.csv", FileMode.Create, FileAccess.Write))
                {
                    csvStream.CopyTo(file);
                }
            }

            Console.WriteLine("VBA certificate information exported to CSV successfully.");
        }
    }
}