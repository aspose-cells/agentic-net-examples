// Title: Hide the Second Worksheet in an Excel Workbook with Aspose.Cells for .NET
// Description: Load an existing Excel file, verify a second worksheet exists, set its IsVisible property to false to hide it, and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells hide worksheet C# | hide second sheet Aspose.Cells | programmatically hide Excel sheet .NET | Workbook.Save after hiding sheet | C# Excel worksheet visibility | Aspose.Cells hide sheet example
// Common Searches: how to hide a specific worksheet with Aspose.Cells | C# hide second worksheet in Excel file | Aspose.Cells hide sheet and save workbook | hide Excel sheet programmatically .NET | make a worksheet invisible using Aspose.Cells
// Developer Intent: Load an existing workbook, hide the second worksheet if it exists, and save the modified file.
// Use Cases: Protect confidential data by hiding internal calculation sheets before distribution. | Create a clean user‑facing workbook where only the primary sheet is visible. | Automate report generation that conceals intermediate worksheets to simplify navigation.
// AI Prompts: Generate C# code with Aspose.Cells that hides the second worksheet of a given workbook and saves it as a new file. | Provide a reusable method that accepts a file path and sheet index, hides the sheet if present, logs a warning for invalid indexes, and returns the output path. | Show how to hide multiple worksheets based on a list of indices using Aspose.Cells for .NET, then save the workbook.

using System;
using Aspose.Cells;

// Load an existing Excel file, verify a second worksheet exists, set its IsVisible property to false to hide it, and save the workbook using Aspose.Cells for C#.
class HideSecondWorksheet
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Ensure there is a second worksheet before attempting to hide it
        if (workbook.Worksheets.Count > 1)
        {
            // Hide the second worksheet (index 1)
            workbook.Worksheets[1].IsVisible = false;
        }

        // Save the workbook with the hidden sheet
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
