// Title: Freeze Header Row and Save as New XLSX with Aspose.Cells for .NET (C#)
// Description: Load an existing workbook, freeze the first row of the first worksheet using FreezePanes, and save the result to a new XLSX file.
// Keywords: Aspose.Cells FreezePanes C# | freeze header row Excel .NET | save workbook new file Aspose | freeze first row Aspose.Cells | C# Excel freeze panes
// Common Searches: Aspose.Cells freeze top row C# | how to freeze header row in Excel using Aspose | save workbook after FreezePanes Aspose.Cells | C# code to freeze first row and export XLSX | Aspose.Cells FreezePanes example
// Developer Intent: Apply a row freeze to the header row of the first worksheet and write the modified workbook to a new XLSX file.
// Use Cases: Create scroll‑friendly reports where column headings stay visible. | Prepare downloadable Excel exports for web apps with locked header rows. | Build reusable templates that automatically freeze the first row before distribution.
// AI Prompts: Generate C# code to freeze the first two rows of a worksheet with Aspose.Cells. | Show how to freeze both the first row and first column, then save the workbook as XLSX. | Explain how to remove frozen panes from a workbook using Aspose.Cells for .NET.

using Aspose.Cells;

// Load an existing workbook, freeze the first row of the first worksheet using FreezePanes, and save the result to a new XLSX file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze the first row (header) by splitting at cell A2
        // This freezes 1 row and 0 columns
        worksheet.FreezePanes("A2", 1, 0);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
