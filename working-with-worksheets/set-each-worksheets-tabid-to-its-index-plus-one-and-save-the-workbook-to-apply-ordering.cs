// Title: Assign sequential TabId values to each worksheet and save the workbook with Aspose.Cells for .NET
// AI Prompts: Iterate through all worksheets, set their TabId to the worksheet index plus one, and save the file using Aspose.Cells in C#. | Reorder Excel sheet tabs by updating the TabId property (1‑based) for each worksheet and write the modified workbook to a new file.
// Common Searches: Aspose.Cells how to change worksheet TabId programmatically in C# | set Excel sheet tab order using TabId property with .NET | C# code to assign sequential TabId to worksheets in an existing workbook | save workbook after modifying TabId values with Aspose.Cells | update Excel tab identifiers for multiple sheets using Aspose.Cells API
// Tags: worksheet TabId assignment Aspose.Cells | Excel sheet tab ordering C# | apply TabId changes Aspose.Cells | save workbook after TabId update .NET | manage sheet tab identifiers Aspose.Cells

using Aspose.Cells;
using System;

// // Loads an Excel workbook, assigns each worksheet a 1‑based TabId to define tab order, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Set each worksheet's TabId to its index plus one
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.TabId = i + 1; // TabId is 1‑based
        }

        // Save the workbook to apply the new ordering
        workbook.Save("output.xlsx");
    }
}
