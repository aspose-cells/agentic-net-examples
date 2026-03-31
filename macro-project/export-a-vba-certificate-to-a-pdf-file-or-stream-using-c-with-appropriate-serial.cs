using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ExportVbaCertificateToPdf
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Path to the workbook that contains a signed VBA project
        string signedWorkbookPath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook sourceWorkbook = new Workbook(signedWorkbookPath);

        // Access the VBA project
        VbaProject vbaProject = sourceWorkbook.VbaProject;

        // Verify that the VBA project is signed
        if (!vbaProject.IsSigned)
        {
            Console.WriteLine("The VBA project is not signed.");
            return;
        }

        // Retrieve the raw certificate data
        byte[] certRawData = vbaProject.CertRawData;
        if (certRawData == null || certRawData.Length == 0)
        {
            Console.WriteLine("Certificate raw data is empty.");
            return;
        }

        // Convert the certificate bytes to a Base64 string for readable representation
        string base64Certificate = Convert.ToBase64String(certRawData);

        // Create a new workbook that will be used to generate the PDF
        Workbook pdfWorkbook = new Workbook();
        Worksheet sheet = pdfWorkbook.Worksheets[0];

        // Write a header and the Base64 certificate into the worksheet
        sheet.Cells["A1"].PutValue("VBA Certificate (Base64):");
        sheet.Cells["A2"].PutValue(base64Certificate);

        // Export the workbook to a PDF file
        string pdfFilePath = "VbaCertificate.pdf";
        pdfWorkbook.Save(pdfFilePath, SaveFormat.Pdf);
        Console.WriteLine($"Certificate exported to PDF file: {pdfFilePath}");

        // Additionally, export the same PDF content to a memory stream
        using (MemoryStream pdfStream = new MemoryStream())
        {
            pdfWorkbook.Save(pdfStream, SaveFormat.Pdf);
            // Optionally, write the stream to another file to verify the output
            File.WriteAllBytes("VbaCertificateFromStream.pdf", pdfStream.ToArray());
            Console.WriteLine("Certificate exported to PDF stream and saved to file: VbaCertificateFromStream.pdf");
        }
    }
}