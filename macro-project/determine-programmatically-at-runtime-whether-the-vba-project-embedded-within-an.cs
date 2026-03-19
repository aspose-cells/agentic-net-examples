using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro-enabled workbook (XLSM)
            string filePath = "sample.xlsm";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Retrieve the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Display the protection status
            Console.WriteLine($"VBA project protected: {isProtected}");
        }
    }
}