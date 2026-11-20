// Title: Copy an Excel theme from a template workbook to a new workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy theme from one workbook to another | how to transfer Excel theme using Aspose.Cells API | clone workbook theme programmatically in .NET | apply template workbook theme to a newly created workbook Aspose.Cells | CopyTheme method example for Excel files in C# | duplicate Excel theme across workbooks with Aspose.Cells
// Tags: CopyTheme Aspose.Cells C# | Excel theme cloning with Aspose.Cells | transfer workbook theme .NET | apply template theme programmatically | Aspose.Cells theme management

// Load the template workbook that contains the desired theme
var templateWorkbook = new Aspose.Cells.Workbook("Template.xlsx");

// Create a new workbook (initially empty)
var newWorkbook = new Aspose.Cells.Workbook();

// Copy the theme from the template workbook to the new workbook
newWorkbook.CopyTheme(templateWorkbook);

// Save the new workbook with the cloned theme
newWorkbook.Save("NewWorkbook.xlsx");
