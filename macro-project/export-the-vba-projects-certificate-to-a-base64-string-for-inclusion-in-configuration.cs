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
                // Path to the input workbook
                string inputPath = "SignedWorkbook.xlsm";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains a signed VBA project
                Workbook workbook = new Workbook(inputPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed and certificate data exists
                if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
                {
                    // Convert the raw certificate bytes to a Base64 string
                    string base64Certificate = Convert.ToBase64String(vbaProject.CertRawData);

                    // Output the Base64 string (could be stored in a configuration file)
                    Console.WriteLine("Base64 encoded VBA certificate:");
                    Console.WriteLine(base64Certificate);

                    // Write the Base64 string to a config file
                    File.WriteAllText("VbaCertificate.config", base64Certificate);
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
                }

                // Save the workbook (no modifications made, but follows lifecycle rule)
                workbook.Save("SignedWorkbook_Processed.xlsm", SaveFormat.Xlsm);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
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