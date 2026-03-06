using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Run()
        {
            // Load the XLTX workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xltx");

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If signed, verify whether the signature is valid
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Signature Valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("No VBA signature to validate.");
            }

            // (Optional) Save the workbook to preserve any changes (none in this case)
            workbook.Save("sample_checked.xltx", SaveFormat.Xltx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureValidationDemo.Run();
        }
    }
}