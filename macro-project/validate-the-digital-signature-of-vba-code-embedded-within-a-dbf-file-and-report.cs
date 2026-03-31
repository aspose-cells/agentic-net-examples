using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateVbaSignatureInDbf
    {
        public static void Run()
        {
            // Load the DBF file into a Workbook object
            Workbook workbook = new Workbook("sample.dbf");

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or not present.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateVbaSignatureInDbf.Run();
        }
    }
}