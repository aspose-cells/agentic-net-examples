// Title: Print cell comments in place for every worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet page setup to print comments in place | Print all comments on each sheet without separate comment pages Aspose.Cells | C# iterate through worksheets and enable PrintInPlace for comments | How to configure PrintComments property for every worksheet in Aspose.Cells | Save workbook after changing PrintComments to PrintInPlace in .NET
// Tags: PrintCommentsType.PrintInPlace configuration | worksheet page setup comment printing | bulk update worksheet print settings C# | Aspose.Cells comment rendering in place | output workbook after print comment changes

// Load an existing workbook
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Set the PrintComments property to print all comments in place for every worksheet
foreach (var sheet in workbook.Worksheets)
{
    sheet.PageSetup.PrintComments = Aspose.Cells.PrintCommentsType.PrintInPlace;
}

// Save the modified workbook
workbook.Save("output.xlsx");
