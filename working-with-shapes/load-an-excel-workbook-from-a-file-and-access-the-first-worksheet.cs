// Title: Load an Excel workbook from a file and get the first worksheet with Aspose.Cells for .NET (C#)
// Description: Shows how to create a Workbook from a local .xlsx file, access the first worksheet via the zero‑based Worksheets collection, and print its Name using Aspose.Cells for .NET.
// Keywords: Aspose.Cells load workbook C# | open Excel file .NET | first worksheet Aspose.Cells | Workbook constructor file path | read worksheet name C#
// Common Searches: Aspose.Cells open existing Excel file C# | C# get first worksheet name Aspose.Cells | load workbook from file Aspose.Cells .NET | how to read worksheet name using Aspose.Cells | Aspose.Cells C# load and access worksheets
// Developer Intent: Open an existing .xlsx file and obtain a reference to its first worksheet.
// Use Cases: Display the first worksheet name in a UI after a user selects an Excel file. | Iterate rows of the first worksheet to import data into a database. | Use the first worksheet as a template for generating a new report with Aspose.Cells.
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, verifies the file exists, and safely returns the first worksheet. | Show how to load a workbook and read all values from the first column of the first worksheet using Aspose.Cells. | Provide an example that catches exceptions when opening an Excel file with Aspose.Cells and logs the worksheet name on success.

using System;
using Aspose.Cells;

// Shows how to create a Workbook from a local .xlsx file, access the first worksheet via the zero‑based Worksheets collection, and print its Name using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet (zero‑based index)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: output the name of the first worksheet
        Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
