using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            ValidateVbaSignatureDemo.Run();
        }
    }

    public class ValidateVbaSignatureDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("SignedVbaWorkbook.xlsm");

            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            if (isSigned)
            {
                bool isValid = workbook.VbaProject.IsValidSigned;
                Console.WriteLine("VBA signature valid: " + isValid);
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no signature to validate.");
            }
        }
    }
}