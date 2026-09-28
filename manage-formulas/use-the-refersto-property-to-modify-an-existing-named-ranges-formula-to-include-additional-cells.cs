// Title: Update a named range's RefersTo address to include extra cells using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, retrieves a specific named range, and expands its RefersTo address to cover additional cells. | Show how to change the RefersTo property of an existing Name object in Aspose.Cells, then save the modified workbook.
// Common Searches: C# Aspose.Cells how to change the address of an existing named range | extend named range RefersTo property programmatically with Aspose.Cells | Aspose.Cells update named range to include additional rows in .NET | modify Excel named range address using Aspose.Cells API
// Tags: Aspose.Cells modify named range RefersTo | C# update Excel named range address | Aspose.Cells extend range programmatically | named range address change Aspose.Cells .NET

using Aspose.Cells;
using System;

// Loads input.xlsx, retrieves the named range "MyRange", updates its RefersTo from the original address to Sheet1!$A$1:$A$10 to expand the range, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range by its name
        Name namedRange = workbook.Worksheets.Names["MyRange"];
        if (namedRange != null)
        {
            // Existing RefersTo string (e.g., Sheet1!$A$1:$A$5)
            string oldRefersTo = namedRange.RefersTo;

            // Define the new range to include additional cells (e.g., extend to A10)
            // Adjust the sheet name and address as needed for your scenario
            string newRefersTo = "Sheet1!$A$1:$A$10";

            // Update the RefersTo property with the new address
            namedRange.RefersTo = newRefersTo;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
