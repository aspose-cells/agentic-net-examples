// Title: Programmatically Hide and Later Unhide an Excel Column Using Aspose.Cells for .NET
// Description: C# sample that creates a workbook, populates columns A‑C, hides column C with Cells.HideColumn, saves the file, then checks a user decision and calls Cells.UnhideColumn to make the column visible before saving a second file. Shows how to toggle column visibility in Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | HideColumn | UnhideColumn | Excel column visibility | programmatic column hide | toggle worksheet column | Excel automation | column visibility control
// Common Searches: Aspose.Cells hide column C example | How to unhide a column with Aspose.Cells after user input | C# hide Excel column programmatically | Toggle column visibility Aspose.Cells .NET | Show hidden column based on condition using Aspose.Cells
// Developer Intent: Hide a worksheet column and reveal it later when a specific condition is satisfied.
// Use Cases: Protect confidential data by hiding a "Secret" column in a report and exposing it after authentication. | Provide a template with certain columns concealed for end‑users, then unhide them when the user selects an "Edit" mode. | Improve performance by initially hiding large data columns and displaying them only when a filter or flag is activated.
// AI Prompts: Generate C# code that uses Aspose.Cells to hide column D, wait for a button click, then unhide the column and restore its original width. | Show how to hide multiple columns (B, C, D) with Aspose.Cells and later unhide them based on a boolean flag while preserving their widths. | Create a reusable method in C# that accepts a column index and a boolean to toggle visibility using Aspose.Cells HideColumn and UnhideColumn.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsColumnToggleDemo
{
    // C# sample that creates a workbook, populates columns A‑C, hides column C with Cells.HideColumn, saves the file, then checks a user decision and calls Cells.UnhideColumn to make the column visible before saving a second file. Shows how to toggle column visibility in Aspose.Cells.
    public class ColumnToggle
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Core demo logic
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data in columns A, B, and C
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Secret");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue("TopSecret1");

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue("TopSecret2");

            // Hide column C (zero‑based index 2)
            cells.HideColumn(2);

            // Save the workbook with the column hidden
            string hiddenPath = "ColumnHidden.xlsx";
            workbook.Save(hiddenPath);
            Console.WriteLine($"Workbook saved with hidden column: {Path.GetFullPath(hiddenPath)}");

            // ----- Later, based on user interaction, decide whether to unhide -----
            bool userWantsToSeeSecret = GetUserDecision();

            if (userWantsToSeeSecret)
            {
                // Unhide column C. Width of -1 uses the standard column width.
                cells.UnhideColumn(2, -1);

                // Save the workbook with the column now visible
                string unhiddenPath = "ColumnUnhidden.xlsx";
                workbook.Save(unhiddenPath);
                Console.WriteLine($"Workbook saved with column unhidden: {Path.GetFullPath(unhiddenPath)}");
            }
        }

        // Mock method representing user interaction; replace with real UI logic.
        private static bool GetUserDecision()
        {
            // For demonstration, we simply return true.
            // In a real scenario, this could be a dialog result, button click, etc.
            return true;
        }
    }
}
