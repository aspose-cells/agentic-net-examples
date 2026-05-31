using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook from disk (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Enumerate all worksheets and output their names to the console
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine(sheet.Name);
        }

        // Save the workbook (save rule) – optional if no changes are made
        workbook.Save("output.xlsx");
    }
}