using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationDemo
    {
        public static void Run()
        {
            // Path to the Excel file that may contain a VBA project
            string filePath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the signature of the VBA project
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature is valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // Optional: Save the workbook (signature state is preserved)
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureValidationDemo.Run();
        }
    }
}