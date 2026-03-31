using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main(string[] args)
    {
        // Verify that a file path was provided
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: CheckVbaProtection <path-to-workbook>");
            return;
        }

        string filePath = args[0];

        // Load the workbook (XLSX, XLSM, etc.)
        Workbook workbook = new Workbook(filePath);

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is password‑protected
        bool isProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine($"Is VBA project protected: {isProtected}");
    }
}