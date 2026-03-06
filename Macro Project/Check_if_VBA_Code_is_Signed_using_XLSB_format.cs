using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureXlsbDemo
    {
        public static void Run()
        {
            // Path to the XLSB workbook
            string filePath = "sample.xlsb";

            // Load the workbook (XLSB format)
            Workbook workbook = new Workbook(filePath);

            // Ensure the workbook contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Determine whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, report whether the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaSignatureXlsbDemo.Run();
        }
    }
}