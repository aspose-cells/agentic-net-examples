// Title: Delete and Verify a Named Range in an Excel Workbook using Aspose.Cells for .NET (C#)
// Description: Loads Input.xlsx, checks for the named range "ObsoleteRange", removes it if present, confirms the deletion, prints counts before and after, and saves the result as Output.xlsx.
// Keywords: Aspose.Cells delete named range | remove named range C# | check named range existence Aspose | verify named range removal | NameCollection Aspose.Cells | C# Excel named range management | .NET workbook cleanup
// Common Searches: Aspose.Cells remove named range C# | how to delete a named range with Aspose.Cells | verify named range deletion Aspose | check if named range exists before removal Aspose.Cells | C# code to delete Excel named range using Aspose
// Developer Intent: Remove the "ObsoleteRange" named range from the workbook and ensure it no longer exists.
// Use Cases: Clean up legacy named ranges before publishing a financial model. | Eliminate temporary ranges created during data preprocessing to avoid naming conflicts. | Confirm successful removal of a range before running dependent calculations or macros.
// AI Prompts: Generate C# code with Aspose.Cells that deletes a specific named range only if it exists and logs the outcome. | Show how to iterate through a workbook's NameCollection to remove multiple named ranges and verify each deletion. | Provide an example of handling a missing named range when calling the Remove method in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads Input.xlsx, checks for the named range "ObsoleteRange", removes it if present, confirms the deletion, prints counts before and after, and saves the result as Output.xlsx.
class DeleteNamedRangeDemo
{
    static void Main()
    {
        // Load an existing workbook that contains the named range "ObsoleteRange"
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the collection of named ranges
        NameCollection names = workbook.Worksheets.Names;

        // Display count before removal
        Console.WriteLine("Named ranges count before removal: " + names.Count);

        // Attempt to remove the named range "ObsoleteRange"
        // The Remove method throws if the name does not exist, so we check first
        bool exists = false;
        foreach (Name n in names)
        {
            if (n.Text == "ObsoleteRange")
            {
                exists = true;
                break;
            }
        }

        if (exists)
        {
            names.Remove("ObsoleteRange");
            Console.WriteLine("Named range \"ObsoleteRange\" removed.");
        }
        else
        {
            Console.WriteLine("Named range \"ObsoleteRange\" not found.");
        }

        // Verify removal by checking the collection again
        bool stillExists = false;
        foreach (Name n in names)
        {
            if (n.Text == "ObsoleteRange")
            {
                stillExists = true;
                break;
            }
        }

        Console.WriteLine("Verification - still exists: " + stillExists);
        Console.WriteLine("Named ranges count after removal: " + names.Count);

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}
