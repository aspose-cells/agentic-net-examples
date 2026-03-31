using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VerifyVbaSignatureDemo
    {
        public static void Run()
        {
            // Path to the workbook that may contain a VBA project
            string workbookFileName = "sample_with_signed_vba.xlsm";
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, workbookFileName);

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Check if the workbook itself is digitally signed (overall file signature)
            Console.WriteLine("Workbook.IsDigitallySigned: " + workbook.IsDigitallySigned);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that a VBA project exists
            if (vbaProject != null)
            {
                // Determine whether the VBA project is signed
                Console.WriteLine("VBA Project IsSigned: " + vbaProject.IsSigned);

                // If signed, verify the validity of the signature
                if (vbaProject.IsSigned)
                {
                    Console.WriteLine("VBA Project IsValidSigned: " + vbaProject.IsValidSigned);
                }
                else
                {
                    Console.WriteLine("VBA Project is not signed; no validity check needed.");
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
            VerifyVbaSignatureDemo.Run();
        }
    }
}