using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaProtectionInOds
    {
        public static void Run()
        {
            // Load the ODS spreadsheet
            Workbook workbook = new Workbook("sample.ods");

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Display whether the VBA project is protected
            Console.WriteLine("VBA Project Protected: " + vbaProject.IsProtected);

            // Display whether the VBA project is signed (as an alternative indicator)
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaProtectionInOds.Run();
        }
    }
}