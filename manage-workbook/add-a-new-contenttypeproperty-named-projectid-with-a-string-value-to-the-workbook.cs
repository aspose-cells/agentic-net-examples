// Title: Add a ProjectId string ContentTypeProperty to an Aspose.Cells workbook in C#
// AI Prompts: Generate C# code that creates a new Workbook, adds a ContentTypeProperty called ProjectId with value "MyProject123", and saves the file as an .xlsx. | Write a C# method that receives a Workbook object and injects a custom string metadata property named ProjectId before saving.
// Common Searches: how to set a custom ContentTypeProperty like ProjectId in an Excel file with Aspose.Cells C# | Aspose.Cells C# add string metadata to workbook contenttypeproperties | example of adding ProjectId property to workbook using Aspose.Cells API | C# code to embed custom workbook metadata in .xlsx with Aspose.Cells | saving Excel workbook with custom ContentTypeProperty using Aspose.Cells
// Tags: add ContentTypeProperty to Excel workbook | Aspose.Cells set custom workbook metadata | C# ContentTypeProperties Add method example | save workbook with ProjectId property | custom string property in Aspose.Cells workbook

using Aspose.Cells;

// // This program creates a new workbook, adds a ContentTypeProperty named "ProjectId" with the value "MyProject123", and saves the workbook as "WorkbookWithProjectId.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a ContentTypeProperty named "ProjectId" with a string value
        workbook.ContentTypeProperties.Add("ProjectId", "MyProject123");

        // Save the workbook to a file
        workbook.Save("WorkbookWithProjectId.xlsx");
    }
}
