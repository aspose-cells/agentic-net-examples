// Title: Aspose.Cells for .NET: Set Sheet Tab Bar Width to 200 px, Hide Third Worksheet, and Save Workbook (C#)
// Description: Creates a new Workbook, sets Settings.SheetTabBarWidth to 2000 (≈200 pixels), ensures at least three worksheets exist, hides the third worksheet (index 2) by setting IsVisible to false, and saves the file as an XLSX document using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# set sheet tab bar width | SheetTabBarWidth 200 pixels | hide worksheet Aspose.Cells | Workbook.Save C# Aspose | Aspose.Cells hide third sheet | adjust tab bar width Aspose | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells set sheet tab bar width 200 | C# hide third worksheet Aspose.Cells | How to change tab bar width in Aspose.Cells | Save workbook after modifying settings Aspose.Cells | Make worksheet invisible Aspose.Cells C#
// Developer Intent: Configure the workbook’s tab bar width to ~200 px, hide the third worksheet, and persist the file.
// Use Cases: Generate Excel files that match a UI design requiring a specific tab bar width. | Store helper or configuration data in a hidden sheet while presenting only visible sheets to end users. | Automate creation of templates where the third sheet is hidden and the tab bar width is standardized. | Prepare workbooks for distribution where corporate UI standards dictate tab bar dimensions.
// AI Prompts: Write C# code with Aspose.Cells to set SheetTabBarWidth to 200 px, hide the worksheet at index 2, and save the workbook as XLSX. | Explain how the SheetTabBarWidth property converts pixel values to the required 1/1000‑window‑width units. | Show how to hide multiple worksheets by index or name in Aspose.Cells while keeping them in the file. | Provide a step‑by‑step guide to ensure a workbook has at least three sheets before hiding one using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new Workbook, sets Settings.SheetTabBarWidth to 2000 (≈200 pixels), ensures at least three worksheets exist, hides the third worksheet (index 2) by setting IsVisible to false, and saves the file as an XLSX document using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the sheet tab bar width.
        // The property expects a value in 1/1000 of the window width.
        // Setting it to 2000 approximates a width of 200 pixels.
        workbook.Settings.SheetTabBarWidth = 2000;

        // Ensure there are at least three worksheets in the workbook.
        while (workbook.Worksheets.Count < 3)
        {
            workbook.Worksheets.Add();
        }

        // Hide the third worksheet (zero‑based index 2).
        workbook.Worksheets[2].IsVisible = false;
        // Alternative method:
        // workbook.Worksheets[2].SetVisible(false, true);

        // Save the workbook to a file.
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
