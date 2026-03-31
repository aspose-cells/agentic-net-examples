using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionStatus
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (macro-enabled workbook). Adjust as needed.
            string filePath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Prepare TSV output
            // Header
            Console.WriteLine("File\tIsProtected\tIsLockedForViewing");

            // Data row
            Console.WriteLine($"{filePath}\t{vbaProject.IsProtected}\t{vbaProject.IslockedForViewing}");
        }
    }
}