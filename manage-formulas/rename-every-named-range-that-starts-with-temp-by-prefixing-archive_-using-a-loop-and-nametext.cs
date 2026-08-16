// Title: Rename "Temp" Named Ranges to "Archive_" with Aspose.Cells for .NET (C#)
// Description: Loads a workbook, iterates through its NameCollection, finds defined names that start with "Temp", prefixes each with "Archive_" using the Name.Text property, and saves the modified file.
// Keywords: Aspose.Cells | C# | .NET | rename named range | NameCollection loop | Name.Text update | prefix Archive_ | temporary named ranges | Excel automation | workbook naming convention
// Common Searches: Aspose.Cells rename named ranges starting with Temp | C# loop through workbook names and add Archive_ prefix | How to change Name.Text for specific Excel names using Aspose | Prefix temporary ranges with Archive_ in .NET | Rename all Temp named ranges programmatically
// Developer Intent: Add an "Archive_" prefix to every defined name that begins with "Temp" in an Excel workbook via Aspose.Cells.
// Use Cases: Archive interim calculation ranges before final report generation. | Enforce a consistent naming scheme for temporary data blocks in shared workbooks. | Prevent naming collisions when versioning a workbook by renaming temporary ranges.
// AI Prompts: Generate C# code using Aspose.Cells that prefixes "Archive_" to all named ranges starting with "Temp" and saves the workbook. | Explain how to safely iterate over a NameCollection and rename entries without triggering collection modification errors. | Provide a logging pattern that records each renamed range and handles null or empty Name.Text values during the process.

using System;
using Aspose.Cells;

namespace RenameTempNamedRanges
{
    // Loads a workbook, iterates through its NameCollection, finds defined names that start with "Temp", prefixes each with "Archive_" using the Name.Text property, and saves the modified file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the collection of defined names in the workbook
            NameCollection names = workbook.Worksheets.Names;

            // Iterate through all names
            for (int i = 0; i < names.Count; i++)
            {
                Name name = names[i];

                // Check if the name starts with "Temp"
                if (name.Text != null && name.Text.StartsWith("Temp", StringComparison.Ordinal))
                {
                    // Prefix with "Archive_"
                    string newName = "Archive_" + name.Text;

                    // Update the name text
                    name.Text = newName;

                    Console.WriteLine($"Renamed '{name.Text}' to '{newName}'");
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
