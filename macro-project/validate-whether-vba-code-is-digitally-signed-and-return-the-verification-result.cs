using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook (provide as first argument or use default)
            string workbookPath = args.Length > 0 ? args[0] : "example.xlsm";

            Workbook workbook;
            bool isSigned = false;
            bool isValidSigned = false;

            if (File.Exists(workbookPath))
            {
                // Load the existing workbook
                workbook = new Workbook(workbookPath);
                if (workbook.VbaProject != null)
                {
                    isSigned = workbook.VbaProject.IsSigned;
                    isValidSigned = workbook.VbaProject.IsValidSigned;
                }
            }
            else
            {
                // Create an empty workbook if the file does not exist
                workbook = new Workbook();
                // No VBA project in a newly created workbook
                isSigned = false;
                isValidSigned = false;
            }

            // Prepare result object
            var verificationResult = new
            {
                IsSigned = isSigned,
                IsValidSigned = isValidSigned
            };

            // Serialize result to JSON
            string jsonResult = JsonSerializer.Serialize(verificationResult);

            // Output JSON
            Console.WriteLine(jsonResult);
        }
    }
}