using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Path to a macro‑enabled workbook that contains a signed VBA project
        string signedWorkbookPath = "SignedWorkbook.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(signedWorkbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed
        if (vbaProject != null && vbaProject.IsSigned)
        {
            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;

            if (certData != null && certData.Length > 0)
            {
                // ----- Export certificate to a file -----
                string certFilePath = "VbaCertificate.cer";
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate saved to file: {certFilePath}");

                // ----- Export certificate to a memory stream -----
                using (MemoryStream certStream = new MemoryStream())
                {
                    certStream.Write(certData, 0, certData.Length);
                    certStream.Position = 0; // Reset for potential reading

                    // Example: read back the first few bytes (optional)
                    byte[] buffer = new byte[10];
                    int read = certStream.Read(buffer, 0, buffer.Length);
                    Console.WriteLine($"Read {read} bytes from certificate stream.");
                }

                // ----- Save the workbook in SpreadsheetML format (Excel 2003 XML) -----
                string xmlOutputPath = "WorkbookExport.xml";
                workbook.Save(xmlOutputPath, SaveFormat.SpreadsheetML);
                Console.WriteLine($"Workbook saved as SpreadsheetML: {xmlOutputPath}");
            }
            else
            {
                Console.WriteLine("The signed VBA project does not contain certificate data.");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a signed VBA project.");
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