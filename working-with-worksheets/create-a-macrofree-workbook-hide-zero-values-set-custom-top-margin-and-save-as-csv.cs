// Title: Create a Macro‑Free Workbook, Hide Zeros, Set Top Margin, Export to CSV with Aspose.Cells (.NET C#)
// Description: Shows how to build a new workbook, strip macros, turn off zero display, adjust the top page margin, and write the file as CSV using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | remove macros workbook | hide zero values | top margin page setup | export to CSV | macro‑free Excel | DisplayZeros property | PageSetup.TopMargin | CSV conversion .NET
// Common Searches: Aspose.Cells remove macros C# | Hide zeros in worksheet Aspose.Cells | Set top margin Aspose.Cells C# | Save workbook as CSV Aspose.Cells | Create macro‑free Excel file with Aspose.Cells
// Developer Intent: Produce a workbook without embedded macros, suppress zero values, define a custom top margin, and save the result as a CSV file using Aspose.Cells in C#.
// Use Cases: Automated generation of clean CSV reports from Excel templates that must not contain macros. | Preparing worksheets with a specific print layout while exporting the underlying data to CSV for downstream processing. | Building a data‑export pipeline that enforces macro removal and consistent page‑setup settings before converting to CSV.
// AI Prompts: Write C# code with Aspose.Cells to create a new workbook, remove any macros, hide zero values, set the top margin to 0.75 inches, and save as CSV. | Explain how the DisplayZeros property and PageSetup.TopMargin affect the output when exporting an Excel sheet to CSV using Aspose.Cells. | Provide a step‑by‑step tutorial for generating a macro‑free workbook, customizing page margins, and converting it to CSV with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to build a new workbook, strip macros, turn off zero display, adjust the top page margin, and write the file as CSV using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure the workbook has no macros
        workbook.RemoveMacro();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Hide zero values in the worksheet
        sheet.DisplayZeros = false;

        // Set a custom top margin (in inches)
        sheet.PageSetup.TopMargin = 0.75; // example: 0.75 inches

        // Save the workbook as a CSV file
        workbook.Save("output.csv", SaveFormat.Csv);
    }
}
