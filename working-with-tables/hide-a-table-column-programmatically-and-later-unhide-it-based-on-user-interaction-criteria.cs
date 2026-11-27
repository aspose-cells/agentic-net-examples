// Title: Hide and later unhide an Excel column with a specific width using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide a column in an Excel worksheet | How to unhide a hidden column and set its width with Aspose.Cells .NET | Toggle Excel column visibility based on user decision using Aspose.Cells | Programmatically hide column B then reveal it later in Aspose.Cells for C#
// Tags: hide column Aspose.Cells .NET | unhide column set width Aspose.Cells | Excel column visibility control Aspose.Cells | conditional column toggle Aspose.Cells C# | programmatic column hide unhide Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, adds sample data, hides column B with Cells.HideColumn, saves the file, then checks a user decision; if true it unhides the column with a defined width using Cells.UnhideColumn and saves a second workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Hide column B (zero‑based index 1)
            sheet.Cells.HideColumn(1);

            // Save the workbook with the column hidden
            string hiddenPath = "HiddenColumn.xlsx";
            workbook.Save(hiddenPath);
            Console.WriteLine($"Workbook saved: {Path.GetFullPath(hiddenPath)}");

            // Determine whether the user wants to unhide the column
            bool unhideRequested = GetUserDecision();

            if (unhideRequested)
            {
                // Unhide column B, specifying a width (e.g., 10 characters)
                sheet.Cells.UnhideColumn(1, 10);

                // Save the workbook with the column visible again
                string unhiddenPath = "UnhiddenColumn.xlsx";
                workbook.Save(unhiddenPath);
                Console.WriteLine($"Workbook saved: {Path.GetFullPath(unhiddenPath)}");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder for user interaction logic
    static bool GetUserDecision()
    {
        // Replace with actual UI or input handling.
        // For demonstration, we simply return true.
        return true;
    }
}
