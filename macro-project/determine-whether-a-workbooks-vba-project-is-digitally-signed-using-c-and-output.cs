using System;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        string filePath = args.Length > 0 ? args[0] : "example.xlsm";

        bool isSigned = false;
        bool isValid = false;

        if (File.Exists(filePath))
        {
            Workbook workbook = new Workbook(filePath);
            if (workbook.VbaProject != null)
            {
                isSigned = workbook.VbaProject.IsSigned;
                isValid = isSigned && workbook.VbaProject.IsValidSigned;
            }
        }

        var result = new
        {
            IsVbaSigned = isSigned,
            IsSignatureValid = isValid
        };

        string json = JsonSerializer.Serialize(result);
        Console.WriteLine(json);
    }
}