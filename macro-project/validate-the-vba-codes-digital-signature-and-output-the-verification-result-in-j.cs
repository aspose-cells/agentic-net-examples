using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Run()
        {
            // Load the workbook that contains a VBA project
            string inputPath = "SignedVbaWorkbook.xlsm";
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Build an anonymous object with the verification results
            var verificationResult = new
            {
                IsSigned = vbaProject.IsSigned,
                IsValidSigned = vbaProject.IsValidSigned
            };

            // Convert the result to JSON
            string json = JsonSerializer.Serialize(verificationResult, new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON string
            Console.WriteLine(json);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            VbaSignatureValidationDemo.Run();
        }
    }
}