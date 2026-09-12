// Title: Hide a column in an Excel worksheet and later unhide it with a specific width using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, hide column B, and save it as HiddenColumn.xlsx with Aspose.Cells. | Load HiddenColumn.xlsx, unhide column B with a width of 10 characters, and save the result as UnhiddenColumn.xlsx using Aspose.Cells. | Write C# code that checks a user‑triggered flag and, if true, unhides a previously hidden column while setting its width via Aspose.Cells.
// Common Searches: Aspose.Cells C# hide column and later unhide with width | how to programmatically hide a column in Excel using Aspose.Cells .NET | unhide hidden column and set column width in Aspose.Cells C# example | conditional column visibility based on user input Aspose.Cells
// Tags: Aspose.Cells hide column example | Aspose.Cells unhide column with width | C# Excel column visibility Aspose.Cells | conditional column show hide Aspose.Cells .NET | set column width on unhide Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, hides column B, saves it, then conditionally loads the file, unhides column B with a width of 10 characters, and saves the updated workbook.
class ColumnVisibilityDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Email");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John Doe");
            sheet.Cells["C2"].PutValue("john.doe@example.com");

            // Hide column B (zero‑based index 1)
            sheet.Cells.HideColumn(1);

            // Save the workbook with the hidden column
            string hiddenPath = "HiddenColumn.xlsx";
            workbook.Save(hiddenPath);

            // -------------------------------------------------
            // Later: based on user interaction criteria, unhide the column
            bool userRequestedUnhide = true; // Simulated UI condition

            if (userRequestedUnhide)
            {
                // Ensure the file exists before loading
                if (!File.Exists(hiddenPath))
                {
                    Console.WriteLine($"File not found: {hiddenPath}");
                    return;
                }

                // Load the workbook that contains the hidden column
                Workbook wbToUnhide = new Workbook(hiddenPath);
                Worksheet ws = wbToUnhide.Worksheets[0];

                // Unhide column B (index 1) and set a default width (e.g., 10 characters)
                ws.Cells.UnhideColumn(1, 10);

                // Save the updated workbook
                string unhiddenPath = "UnhiddenColumn.xlsx";
                wbToUnhide.Save(unhiddenPath);
                Console.WriteLine($"Column B unhidden and saved to {unhiddenPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
