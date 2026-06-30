using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (specify the correct file path)
        LoadOptions loadOptions = new LoadOptions();
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Enumerate all worksheets and output their names to the console
        Console.WriteLine("Worksheet names in the loaded workbook:");
        foreach (Worksheet ws in workbook.Worksheets)
        {
            Console.WriteLine(ws.Name);
        }
    }
}

// Author note: Simple enumeration of worksheet names after loading a workbook.