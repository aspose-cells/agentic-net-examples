using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Path to the Excel workbook that may contain a VBA project
        string inputPath = "sample.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            // Determine if the VBA project is protected
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"VBA Project Protected: {isProtected}");
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }
    }
}