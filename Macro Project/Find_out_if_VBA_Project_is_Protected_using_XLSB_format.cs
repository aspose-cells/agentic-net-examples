using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        // Load the XLSB workbook
        Workbook workbook = new Workbook("sample.xlsb");

        // Get the VBA project (may be null if no VBA project exists)
        VbaProject vbaProject = workbook.VbaProject;

        // Check whether the VBA project is protected
        bool isProtected = vbaProject != null && vbaProject.IsProtected;

        Console.WriteLine("VBA Project Protected: " + isProtected);
    }
}