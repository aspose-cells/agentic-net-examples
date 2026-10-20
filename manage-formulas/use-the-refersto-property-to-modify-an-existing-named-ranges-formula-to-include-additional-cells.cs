// Title: Extend an existing named range by appending cells using the RefersTo property in Aspose.Cells for .NET (C#)
// Common Searches: how to add extra cells to an existing named range with Aspose.Cells C# | using Aspose.Cells RefersTo to concatenate ranges in a workbook | programmatically extend a named range in Excel using Aspose.Cells .NET | append additional rows to a named range RefersTo property in C# | update named range formula to include new cells Aspose.Cells example
// Tags: named range RefersTo manipulation Aspose.Cells | expand Excel named range via code | concatenate multiple ranges C# | modify named range formula .NET | update named range cell addresses

using System;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, retrieves a named range, builds a new RefersTo string that combines the original range with additional cell addresses, assigns the updated RefersTo back to the named range, and saves the workbook.
class ModifyNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the existing named range to modify
            string rangeName = "MyRange";

            // Retrieve the named range object by name
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Existing RefersTo string (e.g., "=Sheet1!$A$1:$A$5")
            string existingRefersTo = namedRange.RefersTo;

            // Define the additional cells to include (adjust sheet name and address as needed)
            string additionalRefersTo = "Sheet1!$A$6:$A$10";

            // Build the new RefersTo expression by appending the additional cells.
            string newRefersTo;
            if (existingRefersTo.StartsWith("="))
            {
                // Remove the leading '=' for concatenation, then prepend it back.
                newRefersTo = "=" + existingRefersTo.Substring(1) + "," + additionalRefersTo;
            }
            else
            {
                newRefersTo = "=" + existingRefersTo + "," + additionalRefersTo;
            }

            // Update the named range's RefersTo property
            namedRange.RefersTo = newRefersTo;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
