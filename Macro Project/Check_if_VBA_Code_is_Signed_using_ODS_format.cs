using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaSignedInOds.Run();
        }
    }

    public class CheckVbaSignedInOds
    {
        public static void Run()
        {
            // Load the ODS workbook
            Workbook workbook = new Workbook("sample.ods");

            // Determine whether the workbook contains any VBA/macros
            if (workbook.HasMacro)
            {
                // Access the VBA project and check its signing status
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("Workbook contains VBA macro.");
                Console.WriteLine("VBA project is signed: " + isSigned);

                // If signed, also report whether the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("Workbook does not contain VBA macro.");
            }
        }
    }
}