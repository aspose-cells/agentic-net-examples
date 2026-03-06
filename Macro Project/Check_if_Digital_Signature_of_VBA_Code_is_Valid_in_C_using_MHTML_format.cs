using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureMhtml
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            string mhtmlPath = "sample.mht";

            if (!File.Exists(mhtmlPath))
            {
                Console.WriteLine($"File not found: {Path.GetFullPath(mhtmlPath)}");
                return;
            }

            // Use Auto load format to let Aspose.Cells detect the file type
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            Workbook workbook = new Workbook(mhtmlPath, loadOptions);

            var vbaProject = workbook.VbaProject;

            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed; no signature to validate.");
            }
        }
    }
}