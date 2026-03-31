using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main()
        {
            string mhtFilePath = "sample.mht";

            if (!File.Exists(mhtFilePath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(mhtFilePath)}");
                return;
            }

            var loadOptions = new LoadOptions(LoadFormat.Auto);
            Workbook workbook = new Workbook(mhtFilePath, loadOptions);

            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Is signature valid? " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}