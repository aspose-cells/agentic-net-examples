using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceLogger
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing macro-enabled workbook (or create a new one if needed)
            // Replace "input.xlsm" with the path to your workbook that contains a VBA project.
            Workbook workbook = new Workbook("input.xlsm");

            // Ensure the workbook has a VBA project before accessing references
            if (workbook.VbaProject != null)
            {
                // Get the collection of VBA project references
                VbaProjectReferenceCollection references = workbook.VbaProject.References;

                // Log each reference name
                for (int i = 0; i < references.Count; i++)
                {
                    VbaProjectReference reference = references[i];
                    Console.WriteLine($"Reference {i + 1}: {reference.Name}");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }

            // Save the workbook if any modifications were made (optional)
            workbook.Save("output.xlsm", SaveFormat.Xlsm);
        }
    }
}