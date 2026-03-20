using System;
using System.Collections.Generic;
using Aspose.Cells;

class RemovePrintTitles
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and clear print title settings
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;

            // Reset rows and columns that are repeated on each printed page
            pageSetup.PrintTitleRows = string.Empty;
            pageSetup.PrintTitleColumns = string.Empty;
        }

        // Remove any defined name that references the previous print titles
        // Excel stores the print‑title range in a built‑in name "_xlnm.Print_Titles"
        NameCollection names = workbook.Worksheets.Names;
        List<string> namesToRemove = new List<string>();

        foreach (Name name in names)
        {
            if (name.Text.Equals("_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase))
            {
                namesToRemove.Add(name.Text);
            }
        }

        if (namesToRemove.Count > 0)
        {
            names.Remove(namesToRemove.ToArray());
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}