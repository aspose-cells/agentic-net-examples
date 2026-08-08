// Title: Load an Excel workbook from a file and get the first worksheet – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing Excel file with Aspose.Cells for .NET by passing the file path to the Workbook constructor, then retrieve the first worksheet (index 0) and output its Name property to the console.
// Keywords: Aspose.Cells load workbook C# | open Excel file Aspose.Cells | first worksheet index 0 | read worksheet name .NET | Aspose.Cells example load file
// Common Searches: Aspose.Cells open existing workbook C# | How to get first worksheet name with Aspose.Cells | C# load Excel file and read sheet name | Aspose.Cells workbook constructor file path
// Developer Intent: Open a saved Excel workbook and obtain a reference to its initial worksheet.
// Use Cases: Verify the workbook structure by displaying the name of the first sheet after loading. | Use the first worksheet as the source for data extraction or transformation. | Chain further operations such as applying formulas, exporting data, or iterating rows on the retrieved sheet.
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells and prints the name of the first worksheet. | Show how to handle missing‑file errors when creating a Workbook from a file path with Aspose.Cells. | Provide an example that accesses the first worksheet and reads the value of cell A1 using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to open an existing Excel file with Aspose.Cells for .NET by passing the file path to the Workbook constructor, then retrieve the first worksheet (index 0) and output its Name property to the console.
public class LoadWorkbookDemo
{
    public static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the file using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet (index 0)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example: display the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
