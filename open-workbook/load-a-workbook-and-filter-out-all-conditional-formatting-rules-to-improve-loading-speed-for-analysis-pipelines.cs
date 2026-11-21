// Title: Remove all conditional formatting from an Excel workbook using Aspose.Cells for .NET to improve load performance
// Common Searches: Aspose.Cells C# clear all conditional formatting to speed up workbook loading | how to delete conditional formatting rules from each worksheet with Aspose.Cells | optimize Excel file load time by removing conditional formats in .NET | C# code to clear ConditionalFormattings collection in Aspose.Cells | remove conditional formatting from large Excel files using Aspose.Cells
// Tags: Aspose.Cells conditional formatting removal | Excel load performance improvement Aspose.Cells | C# clear ConditionalFormattings API | remove worksheet conditional formats .NET | optimize large Excel workbook loading Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook with Aspose.Cells, iterates through every worksheet, clears the ConditionalFormattings collection to delete all conditional formatting rules, and saves the cleaned workbook, resulting in faster load times for downstream analysis.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the collection of conditional formatting rules
            ConditionalFormattingCollection cfCollection = sheet.ConditionalFormattings;

            // Remove all conditional formatting rules to improve loading speed
            cfCollection.Clear();
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
