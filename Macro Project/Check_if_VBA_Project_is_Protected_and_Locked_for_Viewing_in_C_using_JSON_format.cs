using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text.Json;

class CheckVbaProtection
{
    static void Main(string[] args)
    {
        string filePath = args.Length > 0 ? args[0] : "sample.xlsm";

        Workbook workbook = new Workbook(filePath);
        VbaProject vbaProject = workbook.VbaProject;

        bool isProtected = vbaProject != null && vbaProject.IsProtected;

        var result = new
        {
            File = filePath,
            IsProtected = isProtected
        };

        string json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine(json);
    }
}