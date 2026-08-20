// Title: Create a Workbook and Add a PivotData Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook, add a worksheet named "PivotData" for pivot‑table source data, and save the file as PivotWorkbook.xlsx using C#.
// Keywords: Aspose.Cells C# workbook creation | add worksheet for pivot table Aspose.Cells | save workbook as xlsx | initialize pivot data sheet | Aspose.Cells .NET example
// Common Searches: Aspose.Cells add worksheet for pivot table C# | how to create a new workbook with Aspose.Cells | save Aspose.Cells workbook as xlsx | initialize pivot data sheet using Aspose.Cells
// Developer Intent: Create a new workbook and add a dedicated worksheet that will host pivot‑table source data.
// Use Cases: Start an empty workbook before importing or generating pivot‑table data. | Organize source data in a clearly named sheet (e.g., "PivotData") for later pivot‑table creation. | Persist the workbook to disk to confirm the worksheet was added correctly.
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, adds a worksheet called "PivotData", optionally fills it with sample data, and saves the file as PivotWorkbook.xlsx. | Provide an Aspose.Cells .NET example that demonstrates adding a pivot‑data worksheet to a new workbook and exporting the workbook to an XLSX file.

using Aspose.Cells;

// Demonstrates how to instantiate an Aspose.Cells Workbook, add a worksheet named "PivotData" for pivot‑table source data, and save the file as PivotWorkbook.xlsx using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a worksheet that will be used for inserting a pivot table
        Worksheet pivotWorksheet = workbook.Worksheets.Add("PivotData");

        // (Optional) Add sample data to the worksheet here if needed

        // Save the workbook to verify the worksheet was added
        workbook.Save("PivotWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
