using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaProjectSigned
    {
        public static void Main(string[] args)
        {
            // Specify the workbook file name (place the file in the same folder as the executable)
            string fileName = "sample_with_macro.xlsm";

            // Build the full path relative to the current directory
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Ensure the workbook contains a VBA project
            if (!workbook.HasMacro)
            {
                Console.WriteLine("The workbook does not contain any VBA project.");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                // Verify if the signature is valid
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}