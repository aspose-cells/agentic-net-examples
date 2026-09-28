// Title: Hide the horizontal scroll bar, show only the vertical scroll bar in an Aspose.Cells worksheet, then save as XLSX using C#
// AI Prompts: Set the worksheet's ShowHorizontalScrollBar property to false and ShowVerticalScrollBar property to true before saving the workbook. | Create a new Workbook, configure the first worksheet to hide the horizontal scroll bar and display the vertical scroll bar, then save the file to a specified path in XLSX format.
// Common Searches: how to hide horizontal scroll bar in Aspose.Cells C# worksheet | display only vertical scroll bar Aspose.Cells .NET example | Aspose.Cells set scroll bar visibility before saving workbook | C# save workbook as XLSX after customizing worksheet scrollbars Aspose.Cells
// Tags: hide horizontal scroll bar Aspose.Cells | show vertical scroll bar Aspose.Cells | save workbook as xlsx Aspose.Cells C# | worksheet scroll bar visibility Aspose.Cells | Aspose.Cells workbook creation C#

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, accesses the first Worksheet, hides the horizontal scroll bar, shows the vertical scroll bar, and saves the file as output.xlsx while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Save the workbook to disk
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
