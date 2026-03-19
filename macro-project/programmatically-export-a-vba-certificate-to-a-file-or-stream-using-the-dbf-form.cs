using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the macro-enabled workbook that contains a signed VBA project
        string sourcePath = "SignedWorkbook.xlsm";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Export the raw certificate bytes directly to a .cer file
            File.WriteAllBytes("VbaCertificate.cer", vbaProject.CertRawData);

            // Create a new workbook to store the certificate data in DBF format
            Workbook dbfWorkbook = new Workbook();
            Worksheet sheet = dbfWorkbook.Worksheets[0];

            // Convert the binary certificate to a Base64 string for text storage
            string base64Cert = Convert.ToBase64String(vbaProject.CertRawData);

            // Write header and certificate data into cells
            sheet.Cells["A1"].PutValue("CertificateBase64");
            sheet.Cells["B1"].PutValue(base64Cert);

            // Configure DBF save options (save rule)
            DbfSaveOptions dbfOptions = new DbfSaveOptions
            {
                ExportAsString = true // Ensure values are saved as strings
            };

            // Save the workbook as a DBF file using the specified options
            dbfWorkbook.Save("VbaCertificate.dbf", dbfOptions);

            Console.WriteLine("Certificate exported to VbaCertificate.cer and VbaCertificate.dbf.");
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
        }
    }
}