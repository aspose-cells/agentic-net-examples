// Title: C# – Set Active Worksheet to Third Sheet (Zero‑Based Index) with Aspose.Cells
// Description: Shows how to create a workbook, add two extra worksheets, and make the third worksheet (index 2) the active sheet using Aspose.Cells for .NET, then save the file.
// Keywords: Aspose.Cells | C# | ActiveSheetIndex | set active worksheet | zero based index | third worksheet | Workbook example | worksheet selection
// Common Searches: Aspose.Cells set active sheet by index | C# set third worksheet active Aspose | ActiveSheetIndex example | change default sheet Aspose.Cells | select worksheet using zero based index
// Developer Intent: Select the third worksheet as the active sheet in a workbook using its zero‑based index.
// Use Cases: Programmatically open a workbook with the third sheet displayed first. | Switch the active sheet based on a user‑chosen position in a reporting tool. | Prepare a template where the default opened sheet is the third one for presentation. | Automate workbook generation that requires a specific sheet to be active before saving.
// AI Prompts: Generate C# code with Aspose.Cells that sets the active worksheet to index 2 and saves the workbook. | Explain the behavior of ActiveSheetIndex in Aspose.Cells and how to retrieve the active sheet name. | Show how to activate a worksheet by its name instead of its index using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add two extra worksheets, and make the third worksheet (index 2) the active sheet using Aspose.Cells for .NET, then save the file.
class SetActiveWorksheet
{
    static void Main()
    {
        // Create a new workbook (initially contains one worksheet)
        Workbook workbook = new Workbook();

        // Add two more worksheets so the workbook has three sheets
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Set the active worksheet to the third sheet (zero‑based index 2)
        workbook.Worksheets.ActiveSheetIndex = 2;

        // Verify the active sheet name (optional)
        Console.WriteLine("Active Sheet: " + workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex].Name);

        // Save the workbook
        workbook.Save("ActiveThirdSheet.xlsx");
    }
}
