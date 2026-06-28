using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaMetadataDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the macro‑enabled workbook (must contain VBA project)
            string inputPath = "input.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Verify that the workbook actually contains a VBA project
            if (!workbook.HasMacro || workbook.VbaProject == null)
            {
                Console.WriteLine("The specified workbook does not contain a VBA project.");
                return;
            }

            // Extract the VBA project's description.
            // Aspose.Cells does not expose a dedicated Description property,
            // but the Name property typically holds the project description.
            string vbaProjectDescription = workbook.VbaProject.Name;

            // Output the extracted description
            Console.WriteLine($"VBA Project Description (Name): {vbaProjectDescription}");

            // (Optional) Save the workbook if any modifications were made.
            // workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }
}