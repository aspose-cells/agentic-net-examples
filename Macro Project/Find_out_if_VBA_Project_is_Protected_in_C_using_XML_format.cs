using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled workbook (XLSM)
        Workbook workbook = new Workbook("sample.xlsm");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is protected
        bool isProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine($"Is VBA Project Protected: {isProtected}");
    }
}