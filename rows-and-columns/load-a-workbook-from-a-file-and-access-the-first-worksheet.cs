// Title: Load an Excel workbook from a file and retrieve the first worksheet using Aspose.Cells for .NET (C#)
// Description: Shows how to open an existing Excel file (e.g., input.xlsx) with Aspose.Cells, obtain the initial Worksheet object from the Workbook, and print its name to the console.
// Keywords: Aspose.Cells open Excel file C# | Workbook constructor file path | first worksheet access | read worksheet name .NET | load workbook Aspose.Cells | C# Excel file handling
// Common Searches: Aspose.Cells open existing workbook C# | Get first worksheet name Aspose.Cells | C# load Excel file and read sheet name | How to read sheet name after loading workbook Aspose.Cells | Aspose.Cells load workbook from path
// Developer Intent: Open a saved Excel document and obtain a reference to its initial worksheet.
// Use Cases: Log the name of the primary sheet after opening a workbook | Iterate rows and columns of the first sheet for data import | Apply formatting or formulas to the initial worksheet immediately after load | Validate workbook structure by checking the first sheet's name
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells and returns the name of the first worksheet. | Show how to catch and handle FileNotFoundException when creating a Workbook from a path in Aspose.Cells. | Explain the difference between Workbook(filePath) and Workbook(stream) constructors in Aspose.Cells. | Provide an example of accessing the first worksheet and reading its cells after loading a workbook.

using System;
using Aspose.Cells;

// Shows how to open an existing Excel file (e.g., input.xlsx) with Aspose.Cells, obtain the initial Worksheet object from the Workbook, and print its name to the console.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: display the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
