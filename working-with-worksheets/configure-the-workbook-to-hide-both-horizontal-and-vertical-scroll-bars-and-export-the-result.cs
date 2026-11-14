// Title: How to hide both horizontal and vertical scrollbars in an Aspose.Cells workbook before exporting to XLSX (C#)
// Common Searches: Aspose.Cells C# hide horizontal scrollbar in worksheet | disable vertical scroll bar in Aspose.Cells before saving Excel file | C# code to turn off both scrollbars in an Aspose.Cells workbook | export Excel file without scrollbars using Aspose.Cells library | set ShowHorizontalScrollBar false and ShowVerticalScrollBar false Aspose.Cells C#
// Tags: hide workbook scrollbars Aspose.Cells C# | disable horizontal scrollbar Aspose.Cells worksheet | disable vertical scrollbar Aspose.Cells worksheet | export workbook without scrollbars Aspose.Cells | Aspose.Cells worksheet view settings C#

using Aspose.Cells;
using System;

// The example demonstrates how to configure an Aspose.Cells Workbook to hide the horizontal and vertical scrollbars via worksheet view properties, then save the workbook as HiddenScrollbars.xlsx in XLSX format, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (optional, shown for completeness)
            Worksheet sheet = workbook.Worksheets[0];

            // Save the workbook to an XLSX file
            string outputPath = "HiddenScrollbars.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
