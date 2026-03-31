using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Run()
        {
            // Path to the Excel file (XLS) that may contain VBA macros
            string filePath = "sample.xls";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook actually contains a VBA project (macro)
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Determine whether the VBA project is digitally signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project is signed: " + isSigned);

                // If it is signed, optionally check if the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain VBA macros.");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaSignature.Run();
        }
    }
}