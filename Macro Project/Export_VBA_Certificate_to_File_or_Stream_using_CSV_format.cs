using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificateCsv
{
    // Export certificate raw data to a CSV file.
    public static void ExportToFile(string workbookPath, string csvPath)
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project (property)
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the project is signed and data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Convert binary data to a Base64 string for CSV representation
            string base64Data = Convert.ToBase64String(vbaProject.CertRawData);

            // Write the Base64 string as a single CSV line
            File.WriteAllText(csvPath, base64Data, Encoding.UTF8);
        }
    }

    // Export certificate raw data to a CSV stream.
    public static void ExportToStream(string workbookPath, Stream outputStream)
    {
        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the project is signed and data exists
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Convert binary data to a Base64 string for CSV representation
            string base64Data = Convert.ToBase64String(vbaProject.CertRawData);

            // Write the string to the provided stream without closing it
            using (StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8, 1024, leaveOpen: true))
            {
                writer.Write(base64Data);
                writer.Flush();
            }
        }
    }

    static void Main()
    {
        string signedWorkbookPath = "SignedWorkbook.xlsm";
        string csvFilePath = "VbaCertificate.csv";

        // Export to a physical CSV file
        ExportToFile(signedWorkbookPath, csvFilePath);
        Console.WriteLine($"Certificate exported to file: {csvFilePath}");

        // Export to a memory stream (example usage)
        using (MemoryStream memoryStream = new MemoryStream())
        {
            ExportToStream(signedWorkbookPath, memoryStream);
            Console.WriteLine($"Certificate exported to stream, length = {memoryStream.Length} bytes");
        }
    }
}