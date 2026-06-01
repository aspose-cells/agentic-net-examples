using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsDemo
{
    class ExportVbaCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string workbookPath = "SignedWorkbook.xlsm";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Ensure the project is signed and certificate data exists
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    byte[] certData = vbaProject.CertRawData;

                    if (certData != null && certData.Length > 0)
                    {
                        // Export the certificate raw data to a file
                        string outputPath = "VbaCertificate.cer";

                        // Write the certificate bytes to disk
                        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                        {
                            fileStream.Write(certData, 0, certData.Length);
                        }

                        Console.WriteLine($"Certificate exported successfully to '{outputPath}'.");
                    }
                    else
                    {
                        Console.WriteLine("No certificate data found in the VBA project.");
                    }
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a signed VBA project.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
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
}