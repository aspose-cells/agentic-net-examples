using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtectionInWorkbook
{
    static void Main()
    {
        // Load the Excel file that contains a VBA project
        Workbook workbook = new Workbook("sample.xls");

        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            Console.WriteLine("VBA Project IsProtected: " + vbaProject.IsProtected);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }
    }
}