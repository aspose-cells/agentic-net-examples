// Title: Add a custom ContentTypeProperty called ProjectId (string) to an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add custom ContentTypeProperty ProjectId to Excel workbook | How to set string metadata ProjectId in an Excel file using Aspose.Cells .NET | C# example for adding ContentTypeProperties to a workbook with Aspose.Cells | Store custom workbook property ProjectId in .xlsx using Aspose.Cells API | Add custom ContentTypeProperty and save workbook as output.xlsx with Aspose.Cells
// Tags: add ContentTypeProperty Aspose.Cells | custom workbook metadata C# | Aspose.Cells ContentTypeProperties string value | store ProjectId in Excel workbook .NET | Aspose.Cells set custom property Excel file

using Aspose.Cells;
using System;

// Creates a new Workbook, adds a ContentTypeProperty named "ProjectId" with the value "12345", and saves it as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a custom ContentTypeProperty named "ProjectId" with a string value
        workbook.ContentTypeProperties.Add("ProjectId", "12345");

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
