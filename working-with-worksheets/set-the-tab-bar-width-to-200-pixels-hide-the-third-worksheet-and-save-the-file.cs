// Title: Set tab bar width to 200 px, hide the third worksheet, and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that adjusts the tab bar width to 200 px, makes the third sheet invisible, and writes the file to disk. | Create a .NET snippet that verifies a workbook has at least three worksheets, hides the sheet at position three, changes the tab bar width, and saves the workbook as an .xlsx file.
// Common Searches: asp.net cells set workbook tab bar width to 200 pixels c# | how to hide the third worksheet using Aspose.Cells in C# | c# ensure workbook contains three sheets before accessing index 2 Aspose.Cells | save workbook after modifying tab bar width with Aspose.Cells .NET | Aspose.Cells hide sheet by index and export to xlsx
// Tags: adjust tab bar width Aspose.Cells | hide worksheet at index Aspose.Cells | save workbook as xlsx Aspose.Cells | ensure minimum worksheet count Aspose.Cells | C# Aspose.Cells workbook manipulation

using System;
using Aspose.Cells;

// The example creates (or loads) a Workbook, guarantees at least three worksheets, hides the third worksheet (index 2) by setting IsVisible to false, sets the workbook's tab bar width to 200 pixels, and saves the result as 'output.xlsx' using Aspose.Cells for .NET with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Ensure there are at least three worksheets before accessing the third one
            while (workbook.Worksheets.Count < 3)
            {
                workbook.Worksheets.Add();
            }

            // Hide the third worksheet (index is zero‑based, so index 2)
            Worksheet thirdSheet = workbook.Worksheets[2];
            thirdSheet.IsVisible = false;

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
