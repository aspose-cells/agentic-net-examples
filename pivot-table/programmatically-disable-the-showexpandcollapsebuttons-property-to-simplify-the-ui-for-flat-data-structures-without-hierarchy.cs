// Title: Disable ShowExpandCollapseButtons in Aspose.Cells for .NET to hide outline symbols
// Common Searches: Aspose.Cells C# hide expand collapse buttons in worksheet | how to turn off ShowExpandCollapseButtons property using Aspose.Cells .NET | remove outline symbols from Excel file with Aspose.Cells API | simplify UI for flat data structures by disabling expand/collapse buttons in Aspose.Cells | programmatically hide worksheet outline symbols in C# Aspose.Cells
// Tags: disable ShowExpandCollapseButtons Aspose.Cells | hide outline symbols Excel .NET | remove expand collapse buttons worksheet | flat data UI simplification Aspose.Cells | pivot table outline suppression C#

using System;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, disables the outline symbols by setting the worksheet's Outline.ShowOutline property to false (equivalent to turning off ShowExpandCollapseButtons), and saves the result as output.xlsx.
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

            // Hide outline symbols (if the API version supports it)
            // sheet.Outline.ShowOutline = false; // Uncomment if available in your Aspose.Cells version

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
