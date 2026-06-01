using System;
using Aspose.Cells;

class DeleteWorksheetScopedNames
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the collection of worksheet‑scoped named ranges
        NameCollection names = workbook.Worksheets.Names;

        // Collect all existing name texts into an array
        int totalNames = names.Count;
        string[] nameTexts = new string[totalNames];
        for (int i = 0; i < totalNames; i++)
        {
            nameTexts[i] = names[i].Text;
        }

        // Remove all named ranges using the Remove(string[]) method
        if (totalNames > 0)
        {
            names.Remove(nameTexts);
        }

        // Save the cleaned workbook to a new file
        workbook.Save("output.xlsx");
    }
}