// Title: Set BottomMargin to 2 points for every worksheet in an Excel file using Aspose.Cells for .NET to avoid bottom clipping
// AI Prompts: Generate C# code that loads an .xlsx workbook, loops through each worksheet, and sets PageSetup.BottomMargin = 2 (points) with Aspose.Cells. | Show how to programmatically adjust the bottom page margin of all sheets to two points to prevent content from being cut off, then save the updated workbook using Aspose.Cells. | Provide a .NET snippet that applies a 2‑point bottom margin across every worksheet in a workbook and writes the result to a new file.
// Common Searches: aspnet aspocells modify bottom margin for every sheet to two points | c# change Excel page setup bottom margin for all worksheets using Aspose.Cells | prevent bottom text clipping in generated Excel files with Aspose.Cells .NET | set worksheet bottom margin programmatically Aspose.Cells example
// Tags: Aspose.Cells set worksheet bottom margin | C# adjust Excel page margins Aspose.Cells | prevent bottom clipping Excel Aspose.Cells | bulk update page setup margins .NET | bottom margin points Excel file

// Load an existing workbook (replace with your actual file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Set the bottom margin to 2 points for each worksheet to avoid clipping
foreach (var worksheet in workbook.Worksheets)
{
    // BottomMargin is measured in points
    worksheet.PageSetup.BottomMargin = 2;
}

// Save the modified workbook (replace with your desired output path)
workbook.Save("output.xlsx");
