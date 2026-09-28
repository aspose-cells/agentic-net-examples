// Title: Set the active worksheet to the third sheet using a zero‑based index with Aspose.Cells for .NET (C#)
// AI Prompts: Assign workbook.Worksheets.ActiveSheetIndex = 2 in a C# Aspose.Cells workbook and save it as an XLSX file. | Activate the worksheet at zero‑based index 2 in a newly created workbook, verify its name, and export the file.
// Common Searches: Aspose.Cells C# set active worksheet by zero based index example | how to make the third sheet the active sheet in an Aspose.Cells workbook | select worksheet at position 2 in a C# Aspose.Cells workbook | C# Aspose.Cells change active sheet programmatically before saving | activate specific worksheet in Aspose.Cells prior to export
// Tags: ActiveSheetIndex property Aspose.Cells | set active worksheet index C# | activate third worksheet Aspose.Cells | Aspose.Cells workbook active sheet selection | zero based worksheet index C#

using System;
using Aspose.Cells;

namespace AsposeCellsActiveSheetDemo
{
    // The sample creates a Workbook, adds two extra worksheets, sets the ActiveSheetIndex to 2 (the third sheet), prints the active sheet name, and saves the workbook as ActiveSheetSet.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Add two more worksheets so we have at least three sheets
            workbook.Worksheets.Add("SecondSheet");
            workbook.Worksheets.Add("ThirdSheet");

            // Set the active worksheet to the third sheet (zero‑based index 2)
            workbook.Worksheets.ActiveSheetIndex = 2;

            // Optional: verify the active sheet name
            Console.WriteLine("Active Sheet: " + workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex].Name);

            // Save the workbook to a file
            workbook.Save("ActiveSheetSet.xlsx");
        }
    }
}
