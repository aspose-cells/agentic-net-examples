// Title: Add an integer custom document property 'ProjectId' to an existing Excel workbook using Aspose.Cells for .NET
// Common Searches: how to add an integer custom metadata field to an Excel file using Aspose.Cells C# | Aspose.Cells C# set ProjectId as custom property in existing workbook | store custom integer value in Excel workbook with Aspose.Cells .NET | C# code to embed ProjectId metadata into an .xlsx file using Aspose.Cells | save Excel workbook with additional custom properties via Aspose.Cells
// Tags: Aspose.Cells custom document property API | integer custom property in Excel workbook | Workbook.CustomDocumentProperties.Add usage | embed ProjectId metadata in .xlsx | save workbook with custom properties C#

using System;
using Aspose.Cells;

// The example loads 'input.xlsx' into a Workbook, adds a custom document property named 'ProjectId' with the integer value 12345 via the CustomDocumentProperties.Add method, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Create a custom property named "ProjectId" with an integer value
        workbook.CustomDocumentProperties.Add("ProjectId", 12345);

        // Save the workbook (the custom property is now embedded)
        workbook.Save("output.xlsx");
    }
}
