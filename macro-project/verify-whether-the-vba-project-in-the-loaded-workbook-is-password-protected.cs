using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load the workbook that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Determine whether the VBA project is password protected
        bool isProtected = vbaProject.IsProtected;

        // Output the result
        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}