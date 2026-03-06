using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the macro‑enabled workbook that contains a signed VBA project
        string sourcePath = "SignedWorkbook.xlsm";

        // Load the workbook (load rule)
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Access the VBA project
        VbaProject vbaProject = sourceWorkbook.VbaProject;

        // Verify that the VBA project is signed
        if (vbaProject != null && vbaProject.IsSigned)
        {
            // Retrieve the certificate raw data (property rule)
            byte[] certData = vbaProject.CertRawData;

            // Convert the binary certificate to a Base64 string for storage in DBF as text
            string certBase64 = Convert.ToBase64String(certData);

            // Create a new workbook to hold the certificate data (create rule)
            Workbook certWorkbook = new Workbook();
            Worksheet sheet = certWorkbook.Worksheets[0];
            sheet.Name = "Certificate";

            // Write header and certificate data into cells
            sheet.Cells["A1"].PutValue("CertificateBase64");
            sheet.Cells["B1"].PutValue(certBase64);

            // Configure DBF save options to export all values as strings (constructor rule)
            DbfSaveOptions dbfOptions = new DbfSaveOptions
            {
                ExportAsString = true
            };

            // Save the workbook as a DBF file (save rule)
            certWorkbook.Save("VbaCertificate.dbf", dbfOptions);

            // Also demonstrate saving to a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                // Save to stream using the same DBF options (save rule)
                certWorkbook.Save(ms, dbfOptions);
                ms.Position = 0; // Reset stream position for further use

                // Example: write the stream content to another DBF file
                using (FileStream fs = new FileStream("VbaCertificateFromStream.dbf", FileMode.Create, FileAccess.Write))
                {
                    ms.CopyTo(fs);
                }
            }

            Console.WriteLine("VBA certificate exported successfully to DBF file and stream.");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project.");
        }
    }
}