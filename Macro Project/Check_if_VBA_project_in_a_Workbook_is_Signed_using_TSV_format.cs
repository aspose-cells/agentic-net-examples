using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            CheckVbaSignedTSV.Run();
        }
    }

    public static class CheckVbaSignedTSV
    {
        public static void Run()
        {
            // Path to the workbook to be checked
            string inputPath = "example.xlsm";

            // Path where the TSV result will be saved
            string outputPath = "VbaSignatureStatus.tsv";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook (macro‑enabled)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;
            bool isValidSigned = vbaProject.IsValidSigned;

            // Create a TSV line: FileName<TAB>IsSigned<TAB>IsValidSigned
            string tsvLine = $"{Path.GetFileName(inputPath)}\t{isSigned}\t{isValidSigned}";

            // Write the TSV line to the output file (overwrites if exists)
            File.WriteAllText(outputPath, tsvLine);
        }
    }
}