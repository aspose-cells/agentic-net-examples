using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureValidation
{
    public class ValidateVbaSignatureFromMht
    {
        public static void Run(string mhtFilePath)
        {
            var workbook = new Workbook(mhtFilePath);
            var vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Is signature valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the MHT file.");
                return;
            }

            ValidateVbaSignatureFromMht.Run(args[0]);
        }
    }
}