using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = args.Length > 0 ? args[0] : "sample_with_macro.xlsm";

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
                return;
            }

            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}