// Title: Add a ProjectId ContentTypeProperty (string) to an Excel workbook with Aspose.Cells for .NET
// Common Searches: how to add a custom ContentTypeProperty named ProjectId to an Excel file using Aspose.Cells C# | Aspose.Cells set string ContentTypeProperty in workbook programmatically | C# code to store project identifier metadata in Excel workbook with Aspose.Cells | add custom metadata property to .xlsx using Aspose.Cells .NET API | save workbook with ProjectId property using Aspose.Cells
// Tags: add ContentTypeProperty to Excel workbook Aspose.Cells | set ProjectId metadata in .xlsx using C# | Aspose.Cells custom workbook property string | store custom content type property .NET | Excel file custom metadata Aspose.Cells

using Aspose.Cells;

// Creates a new Workbook, adds a string ContentTypeProperty called "ProjectId", and saves the file as "WorkbookWithProjectId.xlsx" using the Aspose.Cells for .NET API.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one if needed)
        Workbook workbook = new Workbook();

        // Add a custom ContentTypeProperty named "ProjectId" with a string value
        workbook.ContentTypeProperties.Add("ProjectId", "YourProjectIdValue");

        // Save the workbook to a file
        workbook.Save("WorkbookWithProjectId.xlsx");
    }
}
