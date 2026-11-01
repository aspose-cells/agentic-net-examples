// Title: Programmatically hide a column in an Aspose.Cells worksheet and later unhide it based on user decision (C# .NET)
// AI Prompts: Generate C# code that uses Aspose.Cells to hide column B, save the workbook, reload it, and conditionally unhide the column with a specific width after evaluating a boolean flag. | Show how to toggle column visibility in an Excel file with Aspose.Cells by calling the hide and unhide methods, including loading the workbook, applying the change, and saving the result.
// Common Searches: how to hide a column in Excel using Aspose.Cells C# | Aspose.Cells unhide column after saving workbook .NET | conditional column visibility based on user input Aspose.Cells | C# example for HideColumn and UnhideColumn with width parameter | reload workbook and change column visibility Aspose.Cells
// Tags: Aspose.Cells hide column functionality C# | Aspose.Cells unhide column with custom width C# | conditional column visibility handling Aspose.Cells .NET | reload workbook and adjust column display Aspose.Cells | Excel column toggle using Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, hides column B, saves it, then reloads the file, checks a simulated user decision, and if true, unhides column B with a width of 10 before saving the final workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            // Hide column B (zero‑based index 1)
            sheet.Cells.HideColumn(1);

            // Save the workbook with the hidden column
            string hiddenFile = "HiddenColumn.xlsx";
            workbook.Save(hiddenFile);

            // ----- Later: decide whether to unhide based on user interaction -----
            bool shouldUnhide = GetUserDecision();

            // Load the previously saved workbook only if it exists
            if (!File.Exists(hiddenFile))
                throw new FileNotFoundException($"The file '{hiddenFile}' was not found.");

            Workbook loadedWorkbook = new Workbook(hiddenFile);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            if (shouldUnhide)
            {
                // Unhide column B; provide a width (e.g., default 10)
                loadedSheet.Cells.UnhideColumn(1, 10);
            }

            // Save the final workbook
            loadedWorkbook.Save("FinalWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simulated method for obtaining user decision.
    // Replace with actual UI or logic as needed.
    static bool GetUserDecision()
    {
        // For demonstration, we simply return true to unhide the column.
        return true;
    }
}
