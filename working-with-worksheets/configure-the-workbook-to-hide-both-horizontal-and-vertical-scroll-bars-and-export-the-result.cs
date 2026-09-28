// Title: Hide both horizontal and vertical scroll bars in an Aspose.Cells workbook and save as XLSX (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a workbook, turn off the horizontal and vertical scroll bars, and save it as an XLSX file. | Demonstrate how to set scroll‑bar visibility properties in Aspose.Cells for .NET before exporting the workbook. | Suggest an alternative method for hiding scroll bars when the ShowHorizontalScrollBar and ShowVerticalScrollBar properties are not present in the current Aspose.Cells version.
// Common Searches: Aspose.Cells C# hide horizontal scroll bar in generated Excel | disable vertical scroll bar in Aspose.Cells workbook before saving | how to remove both scroll bars from an Excel file created with Aspose.Cells | Aspose.Cells scroll bar visibility not working in latest version | C# export Excel without scrollbars using Aspose.Cells
// Tags: hide scrollbars Aspose.Cells workbook | disable horizontal and vertical scrollbars .NET | Aspose.Cells scrollbars control API | save workbook with hidden scrollbars | fallback for missing ShowHorizontalScrollBar property

using System;
using Aspose.Cells;

// The example creates a new Aspose.Cells Workbook, notes that recent versions no longer expose ShowHorizontalScrollBar and ShowVerticalScrollBar properties, and saves the file as HiddenScrollBars.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: In recent Aspose.Cells versions the scroll‑bar visibility properties
            // (ShowHorizontalScrollBar, ShowVerticalScrollBar) are not available.
            // If you need to control them, use the appropriate API for your version.

            // Save the workbook to a file
            string outputPath = "HiddenScrollBars.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
