// Title: Enable printing of gridlines only on the first three worksheets of an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Load a workbook, set PageSetup.PrintGridlines = true for the first three worksheets, and save the updated file. | Iterate through up to three sheets in a .xlsx and turn on gridline printing while leaving all other sheets unchanged using Aspose.Cells C#.
// Common Searches: Aspose.Cells C# enable gridlines on first three worksheets only | Print gridlines for selected sheets in an existing Excel file using Aspose.Cells | Set PrintGridlines property for multiple worksheets in .NET | How to keep default gridline settings on later sheets while changing first sheets in Aspose.Cells
// Tags: Aspose.Cells set PrintGridlines for specific worksheets | C# enable gridline printing on selected Excel sheets | page setup gridlines first three worksheets Aspose.Cells | modify workbook gridline settings Aspose.Cells .NET | selective worksheet printing options Aspose.Cells

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Enable printing of gridlines for the first three worksheets (or fewer if the workbook has less)
int sheetsToProcess = System.Math.Min(3, workbook.Worksheets.Count);
for (int i = 0; i < sheetsToProcess; i++)
{
    // Access the worksheet
    Aspose.Cells.Worksheet sheet = workbook.Worksheets[i];

    // Set the page setup option to print gridlines
    sheet.PageSetup.PrintGridlines = true;
}

// Save the modified workbook
workbook.Save("output.xlsx");
