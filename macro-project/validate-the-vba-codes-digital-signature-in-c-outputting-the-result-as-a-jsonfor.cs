using System;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook("input.xlsm");

            // Access VBA project properties
            bool isSigned = workbook.VbaProject.IsSigned;
            bool isValidSigned = workbook.VbaProject.IsValidSigned;

            // Prepare result object
            var result = new
            {
                IsSigned = isSigned,
                IsValidSigned = isValidSigned
            };

            // Serialize result to JSON
            string jsonResult = JsonSerializer.Serialize(result);

            // Output JSON string
            Console.WriteLine(jsonResult);
        }
    }
}