using System;
using System.Collections.Generic;
using Aspose.Cells;

class HiddenProtectedWorksheetDetector
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // List to hold names of worksheets that are both hidden and protected
        List<string> hiddenProtectedSheets = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine if the worksheet is hidden (VisibilityType not Visible)
            bool isHidden = sheet.VisibilityType != VisibilityType.Visible;

            // Determine if the worksheet is protected
            // A worksheet can be protected without a password (IsProtected)
            // or protected with a password (Protection.IsProtectedWithPassword)
            bool isProtected = sheet.IsProtected || sheet.Protection.IsProtectedWithPassword;

            // If both conditions are true, add the sheet name to the list
            if (isHidden && isProtected)
            {
                hiddenProtectedSheets.Add(sheet.Name);
            }
        }

        // Output the results
        if (hiddenProtectedSheets.Count > 0)
        {
            Console.WriteLine("Hidden and protected worksheets found:");
            foreach (string name in hiddenProtectedSheets)
            {
                Console.WriteLine("- " + name);
            }
        }
        else
        {
            Console.WriteLine("No hidden protected worksheets found.");
        }

        // Save the workbook if needed (no modifications made in this example)
        workbook.Save("output.xlsx");
    }
}