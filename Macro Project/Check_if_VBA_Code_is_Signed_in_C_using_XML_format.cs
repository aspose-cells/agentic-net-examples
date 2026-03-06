using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Run()
        {
            // Load an existing workbook that contains a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Determine whether the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, also verify the validity of the signature
            if (isSigned)
            {
                bool isValid = workbook.VbaProject.IsValidSigned;
                Console.WriteLine("Signature valid: " + isValid);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaSignature.Run();
        }
    }
}