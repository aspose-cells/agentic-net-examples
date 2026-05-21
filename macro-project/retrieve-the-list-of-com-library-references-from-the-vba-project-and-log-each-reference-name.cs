using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class RetrieveVbaReferences
{
    static void Main()
    {
        // Load a macro-enabled workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsm");

        // Check if the workbook contains a VBA project
        if (workbook.VbaProject != null)
        {
            // Get the collection of VBA references
            VbaProjectReferenceCollection references = workbook.VbaProject.References;

            // Iterate through each reference and log its name
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
    }
}