using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Path to the macro-enabled workbook (XLTM)
        string workbookPath = "sample.xltm";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(workbookPath);

        // Retrieve the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Output the protection status
        Console.WriteLine($"Is VBA Project Protected: {isProtected}");
    }
}