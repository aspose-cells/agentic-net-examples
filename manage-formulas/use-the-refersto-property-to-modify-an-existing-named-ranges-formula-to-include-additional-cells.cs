// Title: Append additional cells to an existing named range using the RefersTo property in Aspose.Cells for .NET
// Common Searches: how to add more cells to a named range with Aspose.Cells C# | modify RefersTo of an existing named range in a .xlsx workbook using Aspose.Cells | append extra range to a named range formula Aspose.Cells .NET example | extend named range programmatically in Excel with Aspose.Cells C# | update named range RefersTo string to include new cells Aspose.Cells
// Tags: Aspose.Cells RefersTo property | C# extend named range | update named range formula Aspose.Cells | append range to named range .xlsx | modify workbook named range programmatically

using Aspose.Cells;
using System;

// The example loads an Excel workbook, retrieves a named range by its name, reads its current RefersTo expression, concatenates an additional cell range, assigns the new RefersTo string back to the named range, and saves the workbook with the updated range.
class Program
{
    static void Main()
    {
        // Load the existing workbook.
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range by its name.
        // Replace "MyRange" with the actual name of your range.
        Name namedRange = workbook.Worksheets.Names["MyRange"];
        if (namedRange != null)
        {
            // Current RefersTo string, e.g. "=Sheet1!$A$1:$A$5"
            string currentRefersTo = namedRange.RefersTo;

            // Define the additional cells you want to include.
            // Ensure the sheet name matches the one used in the existing range.
            string additionalRange = ",Sheet1!$A$6:$A$10";

            // Update the RefersTo property to include the new cells.
            // The result will be something like "=Sheet1!$A$1:$A$5,Sheet1!$A$6:$A$10"
            namedRange.RefersTo = currentRefersTo + additionalRange;
        }

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}
