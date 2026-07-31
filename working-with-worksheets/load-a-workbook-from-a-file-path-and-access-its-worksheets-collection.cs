// Title: Load an Excel workbook from a file path and list its worksheets using Aspose.Cells for .NET (C#)
// Description: Shows how to open an existing .xlsx file with Aspose.Cells' Workbook(string) constructor, retrieve the WorksheetCollection via workbook.Worksheets, and iterate through the sheets to output each name. Includes an optional save step.
// Keywords: Aspose.Cells | C# | load workbook | open Excel file | worksheet collection | list sheet names | Workbook(string) constructor | read .xlsx | iterate worksheets | Aspose.Cells .NET
// Common Searches: Aspose.Cells open existing Excel file C# | list worksheet names Aspose.Cells .NET | Workbook constructor string example | get worksheets collection Aspose.Cells | C# read Excel sheet names with Aspose
// Developer Intent: Open an existing Excel workbook and obtain its worksheets collection for further manipulation or analysis.
// Use Cases: Display all sheet names from input.xlsx in a console application. | Verify the presence of a required worksheet before processing data. | Rename or reorder worksheets after loading a workbook based on business rules.
// AI Prompts: Write C# code that loads a workbook from a given path using Aspose.Cells and returns an array of worksheet names. | Provide an example that iterates through the WorksheetCollection and prefixes each sheet name with "Report_". | Explain how to catch FileNotFoundException when opening a workbook with Aspose.Cells and suggest a fallback strategy.

using System;
using Aspose.Cells;

// Shows how to open an existing .xlsx file with Aspose.Cells' Workbook(string) constructor, retrieve the WorksheetCollection via workbook.Worksheets, and iterate through the sheets to output each name. Includes an optional save step.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the specified file path using the Workbook(string) constructor
        Workbook workbook = new Workbook(filePath);

        // Access the collection of worksheets in the loaded workbook
        WorksheetCollection worksheets = workbook.Worksheets;

        // Example: iterate through the worksheets and display their names
        for (int i = 0; i < worksheets.Count; i++)
        {
            Console.WriteLine($"Worksheet {i}: {worksheets[i].Name}");
        }

        // (Optional) Save the workbook to a new file if needed
        // workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
