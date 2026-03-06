using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureMhtDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Resolve the path to the MHT file relative to the executable location
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string mhtPath = Path.Combine(exeDir ?? string.Empty, "sample.mht");

            if (!File.Exists(mhtPath))
            {
                Console.WriteLine($"File not found: {mhtPath}");
                return;
            }

            // Load the workbook from the MHT file (auto-detect format)
            Workbook workbook = new Workbook(mhtPath);

            // Ensure the workbook actually contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Determine whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If it is signed, also check whether the signature is valid
                if (isSigned)
                {
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("Signature valid: " + isValid);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}