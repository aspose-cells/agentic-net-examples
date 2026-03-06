using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignedInXLTM
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Adjust the path to the XLTM (macro‑enabled template) workbook as needed
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "template.xltm");

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the XLTM workbook
            Workbook workbook = new Workbook(filePath);

            // Determine whether the workbook contains any VBA macro
            if (workbook.HasMacro)
            {
                Console.WriteLine("Workbook contains VBA macro.");

                // Check if the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, optionally verify the signature validity
                if (isSigned)
                {
                    Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("Workbook does not contain VBA macro.");
            }
        }
    }
}