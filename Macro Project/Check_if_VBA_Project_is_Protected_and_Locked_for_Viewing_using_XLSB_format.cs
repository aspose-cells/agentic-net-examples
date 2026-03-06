using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB file to be examined
            string filePath = "sample.xlsb";

            // Load the workbook (XLSB format)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Output the protection status of the VBA project
            Console.WriteLine($"Is VBA Project Protected: {vbaProject.IsProtected}");
            Console.WriteLine($"Is VBA Project Locked for Viewing: {vbaProject.IslockedForViewing}");
        }
    }
}