using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ExportVbaCertificate
{
    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Expect two arguments: workbook path and certificate output path.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExportVbaCertificate <workbookPath> <certificateFilePath>");
            return;
        }

        string workbookPath = args[0];
        string certificateFilePath = args[1];

        Run(workbookPath, certificateFilePath);
    }

    // Exports the VBA project's digital certificate to the specified file path.
    public static void Run(string workbookPath, string certificateFilePath)
    {
        try
        {
            // Verify that the workbook file exists before loading.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook that contains the VBA project.
            Workbook workbook = new Workbook(workbookPath);

            // Get the VBA project from the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed.
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data.
                byte[] certData = vbaProject.CertRawData;

                // Ensure the certificate data is present.
                if (certData != null && certData.Length > 0)
                {
                    // Write the certificate data to the specified file.
                    File.WriteAllBytes(certificateFilePath, certData);
                    Console.WriteLine($"Certificate exported successfully to: {certificateFilePath}");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty; nothing was exported.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}