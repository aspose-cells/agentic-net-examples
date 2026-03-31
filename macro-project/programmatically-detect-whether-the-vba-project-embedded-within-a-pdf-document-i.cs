using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class DetectVbaProjectProtection
    {
        static void Main()
        {
            // Path to the workbook that may contain a VBA project
            string filePath = "sample.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project (may be null if none exists)
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is protected
            bool isVbaProtected = vbaProject != null && vbaProject.IsProtected;

            // Output the detection result
            Console.WriteLine($"Is VBA Project Protected: {isVbaProtected}");
        }
    }
}