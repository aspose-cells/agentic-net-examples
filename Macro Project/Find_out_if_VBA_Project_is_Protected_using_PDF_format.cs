using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load a macro-enabled Excel workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Determine whether the VBA project is protected
        bool vbaProtected = workbook.VbaProject.IsProtected;
        Console.WriteLine("Is VBA Project Protected: " + vbaProtected);

        // Save the workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}