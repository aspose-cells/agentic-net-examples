// Title: Copy an Excel worksheet and retain its embedded charts using Aspose.Cells for .NET
// AI Prompts: Write C# code that clones a worksheet and keeps every chart and picture unchanged using Aspose.Cells. | Demonstrate how to add a new sheet, copy all content from an existing sheet, rename it, and save the workbook while preserving visual objects. | Create a reusable method that copies a worksheet and retains embedded chart objects in a .NET application.
// Common Searches: Aspose.Cells copy worksheet keep charts C# | duplicate Excel sheet with embedded objects using .NET library | preserve chart objects when cloning a worksheet Aspose.Cells | C# copy sheet retain images and charts | example of copying a worksheet with all visual elements Aspose.Cells
// Tags: worksheet duplication with chart preservation Aspose.Cells | clone Excel sheet containing embedded objects C# | Aspose.Cells duplicate sheet retain images | copy sheet while keeping visual elements .NET | preserve chart objects during worksheet duplication

// Load the existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Identify the worksheet to be copied (by name or index)
Aspose.Cells.Worksheet sourceSheet = workbook.Worksheets["Sheet1"];   // change name as needed

// Add a new empty worksheet that will receive the copy
int newSheetIndex = workbook.Worksheets.Add();
Aspose.Cells.Worksheet copiedSheet = workbook.Worksheets[newSheetIndex];

// Copy the entire contents of the source worksheet, including all embedded objects (charts, images, etc.)
sourceSheet.Copy(copiedSheet);

// Optionally rename the copied worksheet
copiedSheet.Name = sourceSheet.Name + "_Copy";

// Save the workbook with the copied worksheet
workbook.Save("output.xlsx");
