using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProtectionMht
{
    static void Main()
    {
        string filePath = "sample.mht";

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {Path.GetFullPath(filePath)}");
            return;
        }

        Workbook workbook = new Workbook(filePath);

        Console.WriteLine("Workbook contains VBA: " + workbook.HasMacro);

        VbaProject vbaProject = workbook.VbaProject;

        if (vbaProject != null)
        {
            Console.WriteLine("Is VBA Project Protected: " + vbaProject.IsProtected);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }
    }
}