using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Run(string prnFilePath)
        {
            Workbook workbook = new Workbook(prnFilePath);
            var vbaProject = workbook.VbaProject;

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

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the PRN file.");
                return;
            }

            VbaSignatureValidationDemo.Run(args[0]);
        }
    }
}