using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string workbookPath = "SignedWorkbook.xlsm";

                // Verify the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project exists and is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    // Ensure we have data before writing to file
                    if (certData != null && certData.Length > 0)
                    {
                        string outputFile = "VbaCertificate.cer";

                        // Write the certificate bytes to a file
                        File.WriteAllBytes(outputFile, certData);

                        Console.WriteLine($"Certificate exported successfully to '{outputFile}'. Length: {certData.Length} bytes.");
                    }
                    else
                    {
                        Console.WriteLine("Certificate data is empty.");
                    }
                }
                else
                {
                    Console.WriteLine("VBA project is not signed or not present.");
                }

                // Save the workbook
                workbook.Save("ExportDemo.xlsx");
                Console.WriteLine("Workbook saved as 'ExportDemo.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}