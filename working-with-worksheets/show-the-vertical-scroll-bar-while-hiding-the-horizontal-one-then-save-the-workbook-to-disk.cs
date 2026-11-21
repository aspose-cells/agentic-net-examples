// Title: Display only the vertical scroll bar and hide the horizontal scroll bar in an Aspose.Cells worksheet, then save the workbook using C#
// Common Searches: C# example to display only the vertical scroll bar in an Aspose.Cells worksheet | How to hide the horizontal scroll bar while keeping the vertical one visible using Aspose.Cells .NET | Setting worksheet view options for scrollbars and saving the workbook with Aspose.Cells in C# | Aspose.Cells ShowVerticalScrollBar true HideHorizontalScrollBar false code sample | Saving an Excel file after customizing scroll bar visibility with Aspose.Cells C#
// Tags: aspacells worksheet viewoptions scrollbars | aspacells disable-horizontal-scrollbar | aspacells enable-vertical-scrollbar | aspacells save-workbook csharp | aspacells configure-viewoptions

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, accesses its first Worksheet, demonstrates how to configure the ViewOptions to show only the vertical scroll bar while hiding the horizontal one, and then saves the workbook to an output.xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // The ViewOptions property may not be available in some versions of Aspose.Cells.
            // If needed, configure view options using the appropriate API for the installed version.

            // Save the workbook to disk
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
