using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaProjectSignedDemo
    {
        public static void Run()
        {
            // Load the workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project object exists
            if (vbaProject != null)
            {
                // Check whether the VBA project is signed
                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, optionally check if the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature valid: " + vbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaProjectSignedDemo.Run();
        }
    }
}