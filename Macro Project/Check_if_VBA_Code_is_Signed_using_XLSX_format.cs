using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            string filePath = "sample.xlsm";

            Workbook workbook = new Workbook(filePath);

            if (workbook.HasMacro)
            {
                var vbaProject = workbook.VbaProject;

                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine("VBA project is signed: " + isSigned);

                if (isSigned)
                {
                    Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain any VBA/macros.");
            }
        }
    }
}