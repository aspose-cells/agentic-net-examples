using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

namespace AsposeCellsExamples
{
    public class VbaProjectProtectionCheck
    {
        public static void Run()
        {
            string filePath = "sample.xlsm";
            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;
            bool isProtected = vbaProject != null && vbaProject.IsProtected;
            var result = new { IsVbaProjectProtected = isProtected };
            string json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine(json);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectProtectionCheck.Run();
        }
    }
}