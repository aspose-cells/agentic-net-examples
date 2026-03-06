using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaProjectProtection.Run();
        }
    }

    public class CheckVbaProjectProtection
    {
        public static void Run()
        {
            string filePath = "sample.xlsm";
            Workbook workbook = new Workbook(filePath);
            VbaProject vbaProject = workbook.VbaProject;
            bool isProtected = vbaProject != null && vbaProject.IsProtected;
            Console.WriteLine($"Is VBA Project Protected ({filePath}): {isProtected}");
        }
    }
}