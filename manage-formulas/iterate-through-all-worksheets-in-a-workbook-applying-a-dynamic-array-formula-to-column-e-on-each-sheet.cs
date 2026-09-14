// Title: Apply a SORT dynamic array formula to column E on every worksheet in an Aspose.Cells workbook (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, loops through all worksheets, writes the formula '=SORT(A1:D10)' into cell E1 of each sheet, and saves the result. | Create a C# snippet using Aspose.Cells to add a spill‑range SORT formula starting at E1 on every worksheet of a workbook and persist the changes.
// Common Searches: aspocells c# set the same SORT formula on all worksheets | how to assign a dynamic array formula to column E for every sheet using Aspose.Cells | iterate through worksheets and apply spill formula in C# Aspose.Cells | bulk insert SORT(A1:D10) formula across multiple worksheets Aspose.Cells | c# aspocells apply array formula to each worksheet column E
// Tags: set dynamic array formula Aspose.Cells C# | process each worksheet assign formula | spill SORT formula column E | bulk formula insertion workbook | load save workbook with formulas | apply same formula across sheets

using System;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, defines the dynamic array formula "=SORT(A1:D10)", iterates over every worksheet, places that formula in cell E1 (which spills automatically into column E), and saves the modified workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Define the dynamic array formula (example: sort range A1:D10)
        string dynamicArrayFormula = "=SORT(A1:D10)";

        // Iterate over every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Apply the dynamic array formula to the first cell of column E (E1)
            // The result will automatically spill into the required rows/columns
            sheet.Cells["E1"].Formula = dynamicArrayFormula;
        }

        // Save the workbook with the applied formulas
        workbook.Save("output.xlsx");
    }
}
