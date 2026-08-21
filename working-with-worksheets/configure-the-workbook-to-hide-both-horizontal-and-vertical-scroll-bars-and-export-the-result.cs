// Title: C# – Hide Horizontal & Vertical Scrollbars in an Aspose.Cells Workbook and Export to XLSX
// Description: Creates a new Workbook, disables both the horizontal and vertical scrollbars using Workbook.Settings, and saves the file as HiddenScrollBars.xlsx.
// Keywords: Aspose.Cells hide scrollbars | disable horizontal scrollbar C# | remove vertical scrollbar Aspose.Cells | Workbook.Settings scrollbars | export workbook without scrollbars | Aspose.Cells C# hide scrollbars
// Common Searches: Aspose.Cells hide scrollbars in workbook | C# hide horizontal and vertical scrollbars Aspose.Cells | disable scrollbars before saving Excel with Aspose.Cells | remove scrollbars from generated XLSX using .NET | Aspose.Cells workbook view settings
// Developer Intent: Turn off the horizontal and vertical scrollbars of a workbook and write the result to an XLSX file.
// Use Cases: Embedding a clean Excel view in a web portal where scrollbars are unnecessary. | Generating printable reports that open without UI scrollbars for a polished layout. | Distributing Excel files that need a minimal interface for end‑user consumption.
// AI Prompts: Write C# code with Aspose.Cells to hide both scrollbars, set the workbook to open in full‑screen mode, and save as XLSX. | Explain how to programmatically confirm that scrollbars are hidden in an Aspose.Cells workbook after saving. | Show how to combine scrollbar hiding with other view options (e.g., hide gridlines, set zoom) using Aspose.Cells .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, disables both the horizontal and vertical scrollbars using Workbook.Settings, and saves the file as HiddenScrollBars.xlsx.
class HideScrollBarsDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Hide the horizontal and vertical scroll bars
        workbook.Settings.IsHScrollBarVisible = false;
        workbook.Settings.IsVScrollBarVisible = false;

        // Export the workbook to an XLSX file
        workbook.Save("HiddenScrollBars.xlsx", SaveFormat.Xlsx);
    }
}
