using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ValidateVbaSignatureDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the XLSB workbook that may contain a signed VBA project
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SignedVbaWorkbook.xlsb");

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the signature
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("Is VBA signature valid? " + isSignatureValid);

                // Optionally, display certificate raw data length
                byte[] certData = vbaProject.CertRawData;
                Console.WriteLine("Certificate raw data length: " + (certData?.Length ?? 0));
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}