using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaSignature
{
    static void Main(string[] args)
    {
        string workbookPath = "example.xlsm";

        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        Workbook workbook = new Workbook(workbookPath);
        VbaProject vbaProject = workbook.VbaProject;

        var result = new
        {
            HasMacro = workbook.HasMacro,
            IsSigned = vbaProject?.IsSigned ?? false,
            IsValidSigned = vbaProject?.IsValidSigned ?? false
        };

        string jsonOutput = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine(jsonOutput);
    }
}