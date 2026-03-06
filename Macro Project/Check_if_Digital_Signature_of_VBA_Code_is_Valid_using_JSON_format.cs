using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VbaSignatureJsonCheckDemo
    {
        public static void Run()
        {
            string fileName = "example.xlsm";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook workbook = new Workbook(filePath);
            var vbaProject = workbook.VbaProject;

            var signatureInfo = new
            {
                IsSigned = vbaProject?.IsSigned ?? false,
                IsValidSigned = vbaProject?.IsValidSigned ?? false
            };

            string jsonResult = JsonSerializer.Serialize(signatureInfo, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine(jsonResult);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureJsonCheckDemo.Run();
        }
    }
}