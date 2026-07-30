// Title: C# – Hide Horizontal & Vertical Scroll Bars in Aspose.Cells Workbook and Save as XLSX
// Description: Shows how to use Aspose.Cells WorkbookSettings in C# to disable the horizontal and vertical scroll bars and export the workbook to an XLSX file.
// Keywords: Aspose.Cells hide scrollbars | WorkbookSettings IsHScrollBarVisible false | WorkbookSettings IsVScrollBarVisible false | disable scrollbars C# | Aspose.Cells hide horizontal scrollbar | Aspose.Cells hide vertical scrollbar | export workbook without scrollbars | C# Excel scroll bar visibility | Aspose.Cells .NET hide scrollbars
// Common Searches: Aspose.Cells hide scrollbars C# | How to turn off scroll bars in an Aspose.Cells workbook | Set IsHScrollBarVisible false Aspose.Cells | Remove scrollbars from generated Excel using Aspose.Cells | C# code to hide scrollbars in Excel file
// Developer Intent: Create a workbook, suppress both scroll bars, and save the file.
// Use Cases: Deliver a read‑only Excel template where scrolling is not required. | Generate a printable report that opens without visible scroll bars. | Embed a fixed‑viewport spreadsheet in a web or desktop UI without scroll controls.
// AI Prompts: Provide C# code that hides both horizontal and vertical scroll bars in an existing Aspose.Cells workbook. | Show how to toggle scroll‑bar visibility at runtime using Aspose.Cells WorkbookSettings. | Explain the impact of WorkbookSettings on scroll‑bar display when the file is opened in Excel.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to use Aspose.Cells WorkbookSettings in C# to disable the horizontal and vertical scroll bars and export the workbook to an XLSX file.
    public class HideScrollBarsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access workbook settings
                WorkbookSettings settings = workbook.Settings;

                // Hide the horizontal scroll bar
                settings.IsHScrollBarVisible = false;

                // Hide the vertical scroll bar
                settings.IsVScrollBarVisible = false;

                // Save the workbook with the scroll bars hidden
                string outputPath = "HiddenScrollBars.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved with both horizontal and vertical scroll bars hidden at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideScrollBarsDemo.Run();
        }
    }
}
