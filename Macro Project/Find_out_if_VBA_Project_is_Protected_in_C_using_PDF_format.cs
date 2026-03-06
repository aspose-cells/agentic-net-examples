using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled Excel workbook (XLSM) that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Determine whether the VBA project inside the workbook is protected
        bool vbaProtected = workbook.VbaProject.IsProtected;
        Console.WriteLine("Is VBA Project Protected: " + vbaProtected);

        // Optionally, save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}