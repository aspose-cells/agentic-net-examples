using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            VerifyVbaSignature.Run();
        }
    }

    public class VerifyVbaSignature
    {
        public static void Run()
        {
            // Path to the workbook that may contain a signed VBA project
            string filePath = "SignedVbaWorkbook.xlsm";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify whether the VBA project is signed and whether the signature is valid
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}