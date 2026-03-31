using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Load a macro-enabled workbook that may contain a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed
        if (vbaProject != null && vbaProject.IsSigned)
        {
            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;

            if (certData != null && certData.Length > 0)
            {
                // Convert the binary data to a Base64 string for text representation
                string base64Cert = Convert.ToBase64String(certData);

                // Export the certificate to a text file
                string txtPath = "VbaCertificate.txt";
                File.WriteAllText(txtPath, base64Cert);
                Console.WriteLine($"Certificate exported to text file: {txtPath}");

                // Demonstrate exporting the same data to a MemoryStream for downstream processing
                using (MemoryStream ms = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(ms))
                {
                    writer.Write(base64Cert);
                    writer.Flush();
                    ms.Position = 0; // Reset position for reading

                    // Example: read back the data from the stream
                    using (StreamReader reader = new StreamReader(ms))
                    {
                        string readBack = reader.ReadToEnd();
                        Console.WriteLine($"Read back from stream, length: {readBack.Length}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Certificate raw data is empty.");
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