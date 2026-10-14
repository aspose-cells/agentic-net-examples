// Title: Assign sequential TabId values to each worksheet and save the workbook with Aspose.Cells for .NET
// AI Prompts: Assign a 1‑based TabId to every worksheet in a workbook and persist the changes using Aspose.Cells in C#. | Reorder Excel sheet tabs by setting TabId based on worksheet index and then save the file with Aspose.Cells.
// Common Searches: how to change Excel sheet tab order using TabId with Aspose.Cells C# | C# Aspose.Cells set TabId for each worksheet to reorder tabs | programmatically assign TabId values to worksheets in a .NET workbook
// Tags: set worksheet TabId Aspose.Cells | reorder Excel sheet tabs using TabId .NET | save workbook after modifying TabId C# | assign sequential TabId to worksheets Aspose.Cells

using Aspose.Cells;
using System;

// Loads an Excel file, assigns each worksheet's TabId to its index plus one (1‑based), and saves the workbook to apply the new tab ordering.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Set each worksheet's TabId to its index + 1
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.TabId = i + 1; // TabId is 1‑based
        }

        // Save the workbook to apply the new ordering (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
