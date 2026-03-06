using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationJsonDemo
    {
        public static void Run()
        {
            // Path to the workbook that contains a VBA project
            string workbookPath = "example.xlsm";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Create an anonymous object with the signature information
            var signatureInfo = new
            {
                IsSigned = vbaProject.IsSigned,
                IsValidSigned = vbaProject.IsValidSigned
            };

            // Convert the information to a formatted JSON string
            string jsonResult = JsonSerializer.Serialize(
                signatureInfo,
                new JsonSerializerOptions { WriteIndented = true });

            // Output the JSON to the console
            Console.WriteLine(jsonResult);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureValidationJsonDemo.Run();
        }
    }
}