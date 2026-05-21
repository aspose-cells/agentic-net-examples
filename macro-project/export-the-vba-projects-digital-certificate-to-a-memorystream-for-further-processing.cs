using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ExportVbaCertificateDemo
{
    // Exports the VBA project's digital certificate to a MemoryStream.
    public static MemoryStream ExportCertificate(string workbookPath)
    {
        // Ensure the workbook file exists.
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

        Workbook workbook;
        try
        {
            // Load the workbook that contains the signed VBA project.
            workbook = new Workbook(workbookPath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException("Failed to load workbook.", ex);
        }

        // Access the VBA project.
        VbaProject vbaProject = workbook.VbaProject;
        if (vbaProject == null)
            throw new InvalidOperationException("Workbook does not contain a VBA project.");

        // Verify that the VBA project is signed.
        if (!vbaProject.IsSigned)
            throw new InvalidOperationException("The VBA project is not signed.");

        // Retrieve the raw certificate data.
        byte[] certData = vbaProject.CertRawData;
        if (certData == null || certData.Length == 0)
            throw new InvalidOperationException("Certificate raw data is unavailable.");

        // Write the certificate data into a MemoryStream.
        return new MemoryStream(certData);
    }

    // Demonstrates usage of the ExportCertificate method.
    public static void Run()
    {
        string signedWorkbookPath = "SignedWorkbook.xlsm";

        try
        {
            // Export the certificate to a memory stream.
            using (MemoryStream certificateStream = ExportCertificate(signedWorkbookPath))
            {
                Console.WriteLine($"Certificate exported successfully. Length: {certificateStream.Length}");
                // Further processing of the certificateStream can be performed here.
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Entry point for the application.
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex}");
        }
    }
}