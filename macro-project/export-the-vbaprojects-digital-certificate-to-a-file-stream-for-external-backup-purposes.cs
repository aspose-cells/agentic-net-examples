using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ExportVbaCertificate
{
    public static void Run()
    {
        try
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Verify that the workbook file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate to an in‑memory stream and then to a file
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        memoryStream.Write(certData, 0, certData.Length);
                        memoryStream.Position = 0; // Reset for reading

                        // Write the stream to a file for external backup
                        using (FileStream fileStream = new FileStream("VbaCertificateBackup.cer", FileMode.Create, FileAccess.Write))
                        {
                            memoryStream.CopyTo(fileStream);
                        }
                    }

                    Console.WriteLine("VBA certificate exported successfully.");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed; no certificate to export.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Entry point for the application
public class Program
{
    public static void Main(string[] args)
    {
        ExportVbaCertificate.Run();
    }
}