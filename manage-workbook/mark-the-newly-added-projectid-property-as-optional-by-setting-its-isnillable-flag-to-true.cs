// Title: Make ProjectId Content Type Property Optional (IsNillable = true) in Aspose.Cells for .NET
// Description: Creates a new Workbook, adds a custom content‑type property named "ProjectId", sets its IsNillable flag to true so the field is optional, and saves the file as ProjectIdOptional.xlsx.
// Keywords: Aspose.Cells | .NET | ContentTypeProperty | IsNillable | optional property | ProjectId | Excel metadata | custom workbook property
// Common Searches: Aspose.Cells set custom property optional | How to use IsNillable in Aspose.Cells | Mark Excel metadata field as nillable | Make ProjectId property optional in workbook
// Developer Intent: Set the IsNillable flag of the ProjectId content‑type property to true, making the property optional in the generated Excel file.
// Use Cases: Define optional metadata for templates where ProjectId may be unknown at creation time. | Allow downstream processes to skip ProjectId validation when the value is absent. | Generate reports that include a ProjectId column but permit blank entries for certain rows.
// AI Prompts: Generate C# code using Aspose.Cells to add a "ProjectId" content‑type property and mark it as nillable. | Explain how setting IsNillable to true affects Excel file validation and schema compliance in Aspose.Cells. | Provide a step‑by‑step tutorial for making any custom content‑type property optional in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Properties;

// Creates a new Workbook, adds a custom content‑type property named "ProjectId", sets its IsNillable flag to true so the field is optional, and saves the file as ProjectIdOptional.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add the ProjectId content type property (initially empty)
        workbook.ContentTypeProperties.Add("ProjectId", "", "string");

        // Retrieve the property and mark it as optional (nillable)
        ContentTypeProperty projectIdProp = workbook.ContentTypeProperties["ProjectId"];
        projectIdProp.IsNillable = true;

        // Save the workbook
        workbook.Save("ProjectIdOptional.xlsx");
    }
}
