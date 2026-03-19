using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaProjectSignatureCheck
    {
        public static void Run()
        {
            // Path to the Excel workbook (must be macro-enabled, e.g., .xlsm)
            string workbookPath = "sample.xlsm";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(workbookPath);

            // Ensure the workbook actually contains a VBA project
            if (!workbook.HasMacro)
            {
                Console.WriteLine("The workbook does not contain any VBA project.");
                return;
            }

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is digitally signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, optionally verify the validity of the signature
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("Signature valid: " + isValid);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaProjectSignatureCheck.Run();
        }
    }
}