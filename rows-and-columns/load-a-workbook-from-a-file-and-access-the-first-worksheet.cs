// Title: Load an Excel workbook from file and get the first worksheet with Aspose.Cells for .NET (C#)
// Description: Shows how to instantiate a Workbook from a .xlsx file using Aspose.Cells for .NET, access the first worksheet (index 0), and read its name in C#.
// Keywords: Aspose.Cells | C# | load workbook | open Excel file | first worksheet | Workbook constructor | read worksheet name | .xlsx | Aspose.Cells API
// Common Searches: Aspose.Cells open existing .xlsx C# | How to get first worksheet name using Aspose.Cells | C# load Excel file Aspose.Cells Workbook(string) | Access worksheet by index Aspose.Cells .NET | Read sheet name from workbook Aspose.Cells
// Developer Intent: Open an existing Excel file and retrieve the first worksheet object.
// Use Cases: Display the first worksheet name in a console or UI after loading a workbook. | Use the first sheet as the data source for reporting or transformation. | Apply template formatting or calculations to the first worksheet programmatically. | Export data from the first worksheet to another file format.
// AI Prompts: Write C# code using Aspose.Cells to open 'input.xlsx' and print the name of the first worksheet. | Generate a snippet that loads an Excel workbook with Aspose.Cells and returns the Worksheet at index 0. | Show how to access the first worksheet and iterate its rows using Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Shows how to instantiate a Workbook from a .xlsx file using Aspose.Cells for .NET, access the first worksheet (index 0), and read its name in C#.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "input.xlsx";

            // Load the workbook from the specified file using the Workbook(string) constructor
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet in the workbook (index 0)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Example usage: display the name of the first worksheet
            Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
        }
    }
}
