// Title: Auto-fit Excel column widths to content with Aspose.Cells in C#
// AI Prompts: Write C# code that creates a workbook, populates cells, invokes the column autosizing method, and saves the file. | Show how to automatically adjust column widths after inserting data using Aspose.Cells for .NET. | Demonstrate using Aspose.Cells to size all columns based on their values before exporting the workbook.
// Common Searches: asp.net c# automatically resize Excel columns with Aspose.Cells | example of Worksheet.AutoFitColumns after populating data in Aspose.Cells | c# Aspose.Cells set column width to fit text programmatically | auto adjust column width in generated Excel file using Aspose.Cells .NET
// Tags: Aspose.Cells column width autosizing | C# Excel column size adjustment | fit worksheet columns to content | generate Excel file with adjusted columns | save workbook after column autosizing

using System;
using Aspose.Cells;

// // Creates a workbook, writes sample headers and values, calls AutoFitColumns to size all columns to their content, and saves the result as AutoFitColumns.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue("Short");
        sheet.Cells["B2"].PutValue("A much longer piece of text");

        // Automatically adjust the width of all columns to fit their content
        sheet.AutoFitColumns();

        // Save the workbook to a file
        workbook.Save("AutoFitColumns.xlsx");
    }
}
