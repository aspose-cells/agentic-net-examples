// Title: How to rename a VBA module to DataProcessor by setting its Name property with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# rename VBA module to DataProcessor before saving workbook | set VBA module Name property using Aspose.Cells .NET example | change first VBA module name in Excel file with Aspose.Cells C# code | update VBA project module name programmatically in .xlsx using Aspose.Cells | rename VBA module in workbook without opening Excel using Aspose.Cells
// Tags: rename VBA module Aspose.Cells C# | set VBA module Name property .NET | modify VBA project module name Excel workbook | Aspose.Cells VBA module rename example | update VBA module identifier programmatically

// Load the existing workbook
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Access the VBA project
var vbaProject = workbook.VbaProject;

// Ensure there is at least one VBA module
if (vbaProject != null && vbaProject.Modules.Count > 0)
{
    // Rename the first module (or adjust index as needed)
    var vbaModule = vbaProject.Modules[0];
    vbaModule.Name = "DataProcessor";
}

// Save the workbook with the updated VBA module name
workbook.Save("output.xlsx", Aspose.Cells.SaveFormat.Xlsx);
