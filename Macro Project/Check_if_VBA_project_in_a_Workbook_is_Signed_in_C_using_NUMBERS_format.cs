using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled .xlsm)
            string filePath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook contains a VBA project
            // If there is no VBA project, IsSigned will be false by default
            bool isSigned = workbook.VbaProject.IsSigned;

            // Output result in numeric format: 1 = signed, 0 = not signed
            Console.WriteLine(isSigned ? 1 : 0);
        }
    }
}