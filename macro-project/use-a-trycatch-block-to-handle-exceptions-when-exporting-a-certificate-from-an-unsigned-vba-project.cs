using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        try
        {
            // Path to the workbook that may contain a signed VBA project
            string workbookPath = "UnsignedWorkbook.xlsm";

            // Verify that the file exists before attempting to load it
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
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Ensure data exists before exporting
                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate to a .cer file
                    string outputPath = "ExportedVbaCertificate.cer";
                    File.WriteAllBytes(outputPath, certData);
                    Console.WriteLine($"Certificate exported successfully to {outputPath}.");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed. No certificate to export.");
            }
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            Console.WriteLine($"Error code: {ex.Code}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
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