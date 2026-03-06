using System;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Ensure VbaProject type is available

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaSignatureDemo.Run();
        }
    }

    public class CheckVbaSignatureDemo
    {
        public static void Run()
        {
            // Load the XLSM workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, optionally verify that the signature is valid
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature valid: " + isValid);
            }
        }
    }
}