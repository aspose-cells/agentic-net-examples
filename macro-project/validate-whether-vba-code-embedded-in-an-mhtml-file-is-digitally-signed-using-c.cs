using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateVbaSignatureFromMhtml
{
    static void Main()
    {
        string filePath = Path.Combine(Environment.CurrentDirectory, "sample.mhtml");
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Load the workbook (auto-detect format)
        Workbook workbook = new Workbook(filePath);

        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            Console.WriteLine($"VBA project is signed: {vbaProject.IsSigned}");
            if (vbaProject.IsSigned)
            {
                Console.WriteLine($"VBA signature is valid: {vbaProject.IsValidSigned}");
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}