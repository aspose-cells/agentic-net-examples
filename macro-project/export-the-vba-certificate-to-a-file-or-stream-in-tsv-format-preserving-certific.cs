using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificateToTsv
{
    public static void Run()
    {
        // Path to the workbook that contains a signed VBA project
        string signedWorkbookPath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(signedWorkbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed and certificate data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Retrieve the raw certificate bytes
            byte[] certData = vbaProject.CertRawData;

            // Encode the binary data as Base64 to store safely in a text (TSV) file
            string certBase64 = Convert.ToBase64String(certData);

            // Create a new workbook to hold the TSV data
            Workbook tsvWorkbook = new Workbook();
            Worksheet sheet = tsvWorkbook.Worksheets[0];

            // Write a header and the encoded certificate data
            sheet.Cells["A1"].PutValue("CertificateBase64");
            sheet.Cells["A2"].PutValue(certBase64);

            // Save the workbook as a TSV file
            string tsvFilePath = "VbaCertificate.tsv";
            tsvWorkbook.Save(tsvFilePath, SaveFormat.Tsv);

            // OPTIONAL: also demonstrate saving to a stream
            using (MemoryStream stream = new MemoryStream())
            {
                // Save to stream in TSV format
                tsvWorkbook.Save(stream, SaveFormat.Tsv);

                // Write the stream content to another file for verification
                File.WriteAllBytes("VbaCertificateFromStream.tsv", stream.ToArray());
            }

            Console.WriteLine("Certificate exported successfully to TSV format.");
        }
        else
        {
            Console.WriteLine("VBA project is not signed or certificate data is unavailable.");
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportVbaCertificateToTsv.Run();
    }
}