using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DBF file (replace with actual file path)
            string inputPath = "sample.dbf";

            // Load the DBF file into a Workbook object
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains a VBA project
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
            else
            {
                // Determine whether the VBA project is signed
                bool isSigned = vbaProject.IsSigned;

                Console.WriteLine("VBA project exists.");
                Console.WriteLine("Is Signed: " + isSigned);

                // If signed, optionally display whether the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature Valid: " + vbaProject.IsValidSigned);
                }
            }

            // Optionally save the workbook back to DBF format (no changes made)
            string outputPath = "sample_output.dbf";
            workbook.Save(outputPath, SaveFormat.Dbf);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
    }
}