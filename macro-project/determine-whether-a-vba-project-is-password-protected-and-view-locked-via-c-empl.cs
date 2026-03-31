using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class VbaProtectionCheck
{
    static void Main()
    {
        string workbookPath = "sample.xlsm";

        // Load the workbook (format is auto‑detected)
        Workbook workbook = new Workbook(workbookPath);

        VbaProject vbaProject = workbook.VbaProject;

        bool isProtected = vbaProject.IsProtected;

        Console.WriteLine($"VBA Project Protected: {isProtected}");
    }
}