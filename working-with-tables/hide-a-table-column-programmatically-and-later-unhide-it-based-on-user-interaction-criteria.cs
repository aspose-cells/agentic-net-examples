// Title: C# Example: Hide and Unhide an Excel Column with Aspose.Cells for .NET Based on User Input
// Description: Demonstrates how to create a workbook, add data to column B, hide the column using Cells.HideColumn, save the file, prompt the user, reload the workbook, unhide the column with Cells.UnhideColumn (default width), and save the updated file. Ideal for interactive Excel generation where column visibility is controlled at runtime.
// Keywords: Aspose.Cells | C# | .NET | HideColumn | UnhideColumn | Excel column visibility | programmatic hide column | user driven unhide | workbook save and load | sample code | GitHub example
// Common Searches: hide column in Excel using Aspose.Cells C# | unhide hidden column after saving workbook Aspose.Cells | prompt user to reveal hidden column Aspose.Cells .NET | Aspose.Cells hide column then unhide example | C# code to toggle column visibility in Excel
// Developer Intent: Hide column B in a workbook, persist the change, then optionally unhide it after a user confirms.
// Use Cases: Protect sensitive data by hiding columns until a user authorizes viewing. | Create templates with helper columns that can be revealed on demand. | Temporarily hide processing columns during automation and restore them before distribution.
// AI Prompts: Show how to hide multiple columns and later unhide them conditionally with Aspose.Cells for .NET. | Provide a C# snippet that hides a column, saves the workbook, and reloads it to unhide based on a boolean flag. | Explain how to retain original column widths when unhiding a column using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, add data to column B, hide the column using Cells.HideColumn, save the file, prompt the user, reload the workbook, unhide the column with Cells.UnhideColumn (default width), and save the updated file. Ideal for interactive Excel generation where column visibility is controlled at runtime.
public class HideUnhideColumnDemo
{
    public static void Main()
    {
        // ---------- Create ----------
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data to column B (zero‑based index 1)
        cells["B1"].PutValue("Hidden Column Header");
        cells["B2"].PutValue(123);
        cells["B3"].PutValue(456);

        // Hide column B
        cells.HideColumn(1);

        // ---------- Save ----------
        // Save the workbook after hiding the column
        string hiddenFile = "HiddenColumnDemo.xlsx";
        workbook.Save(hiddenFile);
        Console.WriteLine($"Column B hidden and workbook saved to '{hiddenFile}'.");

        // ---------- User Interaction ----------
        // Ask the user whether to unhide the column
        Console.Write("Do you want to unhide column B? (y/n): ");
        string response = Console.ReadLine();

        if (!string.IsNullOrEmpty(response) && response.Trim().ToLower() == "y")
        {
            // ---------- Load ----------
            // Load the previously saved workbook
            Workbook loadedWorkbook = new Workbook(hiddenFile);
            Cells loadedCells = loadedWorkbook.Worksheets[0].Cells;

            // Unhide column B with standard width (-1 uses default width)
            loadedCells.UnhideColumn(1, -1);

            // ---------- Save ----------
            // Save the workbook after unhiding the column
            string unhiddenFile = "UnhiddenColumnDemo.xlsx";
            loadedWorkbook.Save(unhiddenFile);
            Console.WriteLine($"Column B unhidden and workbook saved to '{unhiddenFile}'.");
        }
        else
        {
            Console.WriteLine("Column B remains hidden.");
        }
    }
}
