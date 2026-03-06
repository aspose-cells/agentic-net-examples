using System;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;

            bool isSigned = vbaProject?.IsSigned ?? false;
            bool isValidSigned = vbaProject?.IsValidSigned ?? false;

            var result = new
            {
                IsSigned = isSigned,
                IsValidSigned = isValidSigned
            };

            string jsonResult = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine(jsonResult);
        }
    }
}