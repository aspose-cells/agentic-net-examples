using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtection
{
    static void Main()
    {
        string filePath = "template.xlsx";

        if (!File.Exists(filePath))
        {
            Workbook wb = new Workbook();
            wb.Save(filePath);
        }

        Workbook workbook = new Workbook(filePath);
        VbaProject vbaProject = workbook.VbaProject;

        bool isProtected = vbaProject != null && vbaProject.IsProtected;

        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}