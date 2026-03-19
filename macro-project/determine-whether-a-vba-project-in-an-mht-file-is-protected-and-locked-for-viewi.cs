using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string mhtFilePath = "sample.mht";

            if (!File.Exists(mhtFilePath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(mhtFilePath)}");
                return;
            }

            // Load the workbook with automatic format detection
            Workbook workbook = new Workbook(mhtFilePath, new LoadOptions(LoadFormat.Auto));

            VbaProject vbaProject = workbook.VbaProject;

            bool isProtected = vbaProject.IsProtected;

            Console.WriteLine($"VBA Project Protected: {isProtected}");
        }
    }
}