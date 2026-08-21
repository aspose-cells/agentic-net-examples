// Title: Aspose.Cells C# – Set Worksheet Zoom to 150 % Programmatically
// Description: Creates a workbook, accesses the first worksheet, assigns the Zoom property a value of 150 to display the sheet at 150 % scale, writes the zoom level to the console, and saves the file as ZoomedWorksheet.xlsx.
// Keywords: Aspose.Cells | C# | worksheet zoom | set zoom 150 | Excel view scale | Zoom property | programmatic Excel zoom | Aspose.Cells workbook
// Common Searches: set worksheet zoom Aspose.Cells C# | Aspose.Cells change Excel sheet zoom programmatically | C# set zoom factor on worksheet | increase Excel view zoom with Aspose.Cells | how to set zoom level 150% in generated Excel file
// Developer Intent: Apply a 150 % zoom to a worksheet using Aspose.Cells before saving the workbook.
// Use Cases: Generate a new Excel file where the default sheet opens at 150 % for detailed inspection. | Open an existing workbook, adjust a specific sheet’s Zoom property to 150 %, and persist the change. | Confirm the applied zoom by outputting worksheet.Zoom to the console after modification.
// AI Prompts: Write C# code with Aspose.Cells that sets the zoom level of a given worksheet to 150 % and saves the workbook. | Show how to assign different zoom percentages to multiple worksheets in the same workbook using Aspose.Cells. | Explain how to read, modify, and verify the Zoom property of a worksheet after loading an existing Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, accesses the first worksheet, assigns the Zoom property a value of 150 to display the sheet at 150 % scale, writes the zoom level to the console, and saves the file as ZoomedWorksheet.xlsx.
class SetWorksheetZoom
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the zoom factor to 150%
        worksheet.Zoom = 150;

        // Output the current zoom level for verification
        Console.WriteLine($"Worksheet zoom set to {worksheet.Zoom}%");

        // Save the workbook to a file
        workbook.Save("ZoomedWorksheet.xlsx");
    }
}
