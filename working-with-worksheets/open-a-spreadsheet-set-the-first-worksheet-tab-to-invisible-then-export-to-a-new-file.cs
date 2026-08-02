// Title: Hide First Worksheet Tab and Export Workbook – Aspose.Cells C# Example
// Description: Load an existing Excel file with Aspose.Cells for .NET, set the first worksheet's visibility to hidden, and save the workbook as a new file while preserving the hidden state. Ideal for creating reports or templates with concealed sheets.
// Keywords: Aspose.Cells C# hide worksheet | Excel hide sheet tab .NET | set worksheet visibility false | save workbook new file Aspose | hidden worksheet export | Aspose.Cells example GitHub | C# Excel API hide tab
// Common Searches: Aspose.Cells hide first worksheet tab | C# hide Excel sheet and save new file | how to make a worksheet invisible with Aspose.Cells | export workbook with hidden sheet .NET | Aspose.Cells hide tab without deleting data
// Developer Intent: Hide the first worksheet tab and save the workbook as a separate file.
// Use Cases: Distribute a client‑facing report while keeping internal calculation sheets hidden. | Provide end users with a template that contains a concealed configuration sheet. | Prepare a presentation workbook where the summary sheet is hidden from view.
// AI Prompts: Generate C# code using Aspose.Cells to hide multiple worksheets by name and save the workbook. | Explain the difference between Worksheet.IsVisible = false and Worksheet.VisibilityType = VisibilityType.Hidden in Aspose.Cells. | Show how to programmatically unhide a hidden worksheet after the file has been saved.

using System;
using Aspose.Cells;

// Load an existing Excel file with Aspose.Cells for .NET, set the first worksheet's visibility to hidden, and save the workbook as a new file while preserving the hidden state. Ideal for creating reports or templates with concealed sheets.
class Program
{
    static void Main()
    {
        // Load an existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Hide the first worksheet tab
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.IsVisible = false; // alternatively: firstSheet.VisibilityType = VisibilityType.Hidden;

        // Save the workbook to a new file (the hidden sheet remains hidden)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
