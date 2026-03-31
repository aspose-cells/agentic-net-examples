using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaProjectSignatureCheck
    {
        public static void Run()
        {
            // Path to the XLTX workbook (template) to be examined
            string workbookPath = "TemplateWithVba.xltx";

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(workbookPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project; it may be null if the workbook has no VBA project
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Determine whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, check whether the signature is valid
            if (isSigned)
            {
                bool isValidSigned = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValidSigned);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed, so no signature validity to report.");
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