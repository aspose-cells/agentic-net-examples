// Title: Clone an Excel workbook and ensure all named ranges are retained using Aspose.Cells for .NET
// Common Searches: Aspose.Cells copy workbook preserving named ranges C# | how to verify named ranges after cloning an Excel file with Aspose.Cells | C# code to deep copy an .xlsx workbook and check named range integrity | compare named range collections between original and cloned workbook Aspose.Cells | validate that named ranges are unchanged after workbook.Copy in .NET
// Tags: workbook copy preserving named ranges Aspose.Cells | named range integrity verification after workbook clone | deep copy Excel workbook using Aspose.Cells Copy method | compare named range collections C# Aspose.Cells | validate named range references in cloned .xlsx file

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, creates a deep copy of its workbook with Aspose.Cells' Copy method, then iterates through the original and cloned worksheets' named range collections to confirm that each name and its reference are identical before saving the cloned workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string originalPath = "original.xlsx";
            const string clonePath = "clone.xlsx";

            // Verify the source file exists to avoid FileNotFoundException
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Error: File \"{originalPath}\" not found.");
                return;
            }

            // Load the original workbook containing named ranges
            Workbook originalWb = new Workbook(originalPath);

            // Create a new workbook and copy the original workbook into it (deep copy)
            Workbook clonedWb = new Workbook();
            clonedWb.Copy(originalWb);

            // Verify that all named ranges are preserved in the clone
            bool allPreserved = true;

            // Get the named ranges collections from both workbooks
            var originalNames = originalWb.Worksheets.Names;
            var clonedNames = clonedWb.Worksheets.Names;

            // Compare the number of named ranges
            if (originalNames.Count != clonedNames.Count)
            {
                allPreserved = false;
            }
            else
            {
                // Compare each named range's name and reference
                for (int i = 0; i < originalNames.Count; i++)
                {
                    var origName = originalNames[i];
                    var cloneName = clonedNames[i];

                    if (origName.Text != cloneName.Text || origName.RefersTo != cloneName.RefersTo)
                    {
                        allPreserved = false;
                        break;
                    }
                }
            }

            // Output verification result
            Console.WriteLine(allPreserved
                ? "All named ranges are preserved in the cloned workbook."
                : "Named ranges mismatch after cloning.");

            // Save the cloned workbook
            clonedWb.Save(clonePath);
            Console.WriteLine($"Cloned workbook saved as \"{clonePath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
