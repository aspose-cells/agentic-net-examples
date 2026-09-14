// Title: How to clear a worksheet’s background image in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets the worksheet’s BackgroundImage property to null, and saves the updated workbook. | Show a step‑by‑step example of removing a sheet’s background picture programmatically in Aspose.Cells without affecting other sheet data.
// Common Searches: asp.net aspose.cells remove background picture from a specific worksheet | c# code to delete worksheet background image in an existing Excel file | how to set worksheet BackgroundImage to null using Aspose.Cells library | example of clearing sheet background image before saving workbook with Aspose.Cells | programmatically remove Excel sheet background graphic in .NET
// Tags: remove worksheet background graphic Aspose.Cells | set BackgroundImage property null C# | Aspose.Cells delete sheet background picture | save workbook after removing background image | excel background image removal .NET

using Aspose.Cells;
using System.IO;

// Loads an Excel workbook, accesses the first worksheet, clears its background image by assigning null to the BackgroundImage property, and saves the modified file as a new workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet you want to modify (e.g., the first worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Clear the worksheet's background image by assigning null to the BackgroundImage stream
        sheet.BackgroundImage = null;

        // Save the changes to a new file
        workbook.Save("output.xlsx");
    }
}
