using System;
using Aspose.Cells;

namespace AsposeCellsSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be examined
            string filePath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Determine whether the workbook is digitally signed
            bool isSigned = workbook.IsDigitallySigned;

            // Output the result
            Console.WriteLine("Workbook is digitally signed: " + isSigned);

            // Clean up resources
            workbook.Dispose();
        }
    }
}