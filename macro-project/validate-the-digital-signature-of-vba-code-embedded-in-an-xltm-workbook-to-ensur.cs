using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            ValidateVbaSignature.Run();
        }
    }

    public static class ValidateVbaSignature
    {
        public static void Run()
        {
            string fileName = "SignedVbaWorkbook.xltm";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature is valid: " + isSignatureValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}