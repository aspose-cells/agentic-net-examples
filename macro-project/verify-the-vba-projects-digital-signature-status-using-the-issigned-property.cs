using System;
using System.IO;
using Aspose.Cells;

namespace AspNetExamples
{
    public class VbaSignatureVerifier
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the macro-enabled workbook (adjust as needed)
            string workbookPath = "sample.xlsm";

            // Verify that the file exists before attempting to load
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found - {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the VBA project from the workbook
                var vbaProject = workbook.VbaProject;

                // Verify if the VBA project is signed and output the result
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    Console.WriteLine("VBA project is signed.");
                    Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("VBA project is not signed.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}