// Title: Hide Excel error values on print with Aspose.Cells for .NET (PrintErrorsBlank)
// Description: Demonstrates how to set the PageSetup.PrintErrors property to PrintErrorsType.PrintErrorsBlank so that error cells (e.g., #DIV/0!, #N/A) appear blank when a worksheet is printed, then saves the workbook as PrintErrorsBlank.xlsx.
// Keywords: Aspose.Cells PrintErrorsBlank | PrintErrors property .NET | hide Excel errors on print | suppress #DIV/0! in printed workbook | Aspose.Cells PageSetup | C# Excel printing errors | Aspose.Cells documentation example
// Common Searches: Aspose.Cells hide error values when printing | Set PrintErrors to blank in C# Aspose.Cells | PrintErrorsBlank example Aspose.Cells .NET | Remove #N/A from printed Excel using Aspose | How to suppress Excel error cells on print with Aspose
// Developer Intent: Configure a worksheet’s PageSetup.PrintErrors to PrintErrorsBlank so that any error values are rendered as empty cells during printing.
// Use Cases: Create printable financial statements that hide calculation errors. | Generate client‑facing invoices where error symbols must not appear on paper. | Prepare archived reports that require a clean printout without Excel error markers.
// AI Prompts: Provide C# code to apply PrintErrorsBlank to all worksheets in an Aspose.Cells workbook and save the file. | Explain the effect of PrintErrorsType.PrintErrorsBlank on printed Excel output and how to switch back to showing errors. | Show how to hide error values only on selected sheets while leaving other sheets unchanged using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to set the PageSetup.PrintErrors property to PrintErrorsType.PrintErrorsBlank so that error cells (e.g., #DIV/0!, #N/A) appear blank when a worksheet is printed, then saves the workbook as PrintErrorsBlank.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Suppress error values during printing by displaying them as blank
        sheet.PageSetup.PrintErrors = PrintErrorsType.PrintErrorsBlank;

        // Save the workbook to a file
        workbook.Save("PrintErrorsBlank.xlsx");
    }
}
