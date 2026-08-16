// Title: List COM Library References in a VBA Project of an XLSM Workbook with Aspose.Cells for .NET
// Description: Loads a macro‑enabled workbook, verifies the presence of a VBA project, accesses its VbaProjectReferenceCollection, and prints each reference name to the console.
// Keywords: Aspose.Cells | VBA references | COM libraries | XLSM | C# | .NET | enumerate VBA references | Workbook.VbaProject | extract reference names | Excel macro automation
// Common Searches: Aspose.Cells get VBA reference names | C# list COM references in .xlsm | how to read VBA project references with Aspose | enumerate VBA libraries in Excel workbook .NET | retrieve VBA reference collection Aspose.Cells
// Developer Intent: Obtain the names of every COM reference defined in the VBA project of a macro‑enabled Excel file.
// Use Cases: Confirm required COM libraries are available before executing macros. | Create a reference inventory for documentation or audit purposes. | Identify missing or broken references across a batch of macro workbooks.
// AI Prompts: Generate C# code using Aspose.Cells that extracts all VBA reference names from a workbook and writes them to a text file. | Show how to filter the VBA references by GUID and log only those matching a specific pattern. | Provide an example that gracefully handles workbooks without a VBA project and logs an appropriate message.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro‑enabled workbook, verifies the presence of a VBA project, accesses its VbaProjectReferenceCollection, and prints each reference name to the console.
class Program
{
    static void Main()
    {
        // Load a macro-enabled workbook that contains a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Check if the workbook has a VBA project
        if (workbook.VbaProject != null)
        {
            // Get the collection of VBA references
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
    }
}
