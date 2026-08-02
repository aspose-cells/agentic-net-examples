// Title: Select a Worksheet by Index or Name with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to retrieve a specific worksheet from an Aspose.Cells workbook using the Worksheets collection—either by its zero‑based index or by its sheet name—then writes data to the selected sheets and saves the workbook.
// Keywords: Aspose.Cells select worksheet C# | Workbook.Worksheets index | Workbook.Worksheets name | .NET spreadsheet API | C# get worksheet by name | C# get worksheet by index | Aspose.Cells example
// Common Searches: Aspose.Cells get worksheet by index C# | Aspose.Cells access worksheet by name | How to select a sheet in Aspose.Cells .NET | Retrieve specific worksheet Aspose.Cells | C# Aspose.Cells worksheet selection example
// Developer Intent: Retrieve a particular worksheet from a workbook using either its index or its name.
// Use Cases: Write data to a known sheet without searching by name. | Apply formatting, formulas, or protection to a sheet identified by its tab name. | Validate worksheet properties (Name, Index) before performing further operations.
// AI Prompts: Show C# code that selects a worksheet by index and writes a value using Aspose.Cells. | Provide an Aspose.Cells example that accesses a worksheet by name and changes its tab color in .NET. | Explain how to loop through workbook.Worksheets to locate a sheet by name and then protect it with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to retrieve a specific worksheet from an Aspose.Cells workbook using the Worksheets collection—either by its zero‑based index or by its sheet name—then writes data to the selected sheets and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Rename the default first worksheet
        workbook.Worksheets[0].Name = "FirstSheet";

        // Add additional worksheets with specific names
        Worksheet secondSheet = workbook.Worksheets.Add("SecondSheet");
        Worksheet thirdSheet = workbook.Worksheets.Add("ThirdSheet");

        // Select a worksheet by its zero‑based index
        Worksheet sheetByIndex = workbook.Worksheets[1]; // This is "SecondSheet"

        // Select a worksheet by its name
        Worksheet sheetByName = workbook.Worksheets["ThirdSheet"]; // This is "ThirdSheet"

        // Verify the selections
        Console.WriteLine($"Selected by index: Name={sheetByIndex.Name}, Index={sheetByIndex.Index}");
        Console.WriteLine($"Selected by name:  Name={sheetByName.Name}, Index={sheetByName.Index}");

        // Write some data to the selected worksheets
        sheetByIndex.Cells["A1"].PutValue("Accessed via index");
        sheetByName.Cells["A1"].PutValue("Accessed via name");

        // Save the workbook to a file
        workbook.Save("SelectWorksheetDemo.xlsx");
    }
}
