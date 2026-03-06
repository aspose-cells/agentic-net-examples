using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        string fileName = "TemplateWithVba.xltm";
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook = new Workbook(filePath);

        bool isSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA project is signed: " + isSigned);

        if (isSigned)
        {
            Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
        }
    }
}