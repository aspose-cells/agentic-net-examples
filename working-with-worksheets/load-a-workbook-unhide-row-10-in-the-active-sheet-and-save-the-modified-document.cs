// Title: Unhide row 10 in the active worksheet and save the workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# code to unhide a hidden row in the active sheet | C# example for changing row visibility and saving an Excel file using Aspose.Cells | How to programmatically reveal row 10 in an existing .xlsx and write to a new file with Aspose.Cells | Load workbook, modify row IsHidden property, and save output using Aspose.Cells .NET | Unhide specific row in active worksheet and export workbook in C#
// Tags: row visibility manipulation Aspose.Cells C# | active worksheet unhide operation .NET | load workbook and modify IsHidden property Aspose.Cells | save modified Excel file Aspose.Cells C# | file existence check before loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, retrieves the active worksheet, sets the IsHidden property of row 10 (zero‑based index 9) to false, and saves the updated workbook as output.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook from the existing file
            var workbook = new Workbook(inputPath);

            // Get the active worksheet
            int activeIndex = workbook.Worksheets.ActiveSheetIndex;
            var sheet = workbook.Worksheets[activeIndex];

            // Unhide row 10 (zero‑based index 9) via the Cells.Rows collection
            sheet.Cells.Rows[9].IsHidden = false;

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
