// Title: Select a Worksheet by Name and Access Its Cells with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a sheet called "Report", retrieve that sheet using the name indexer, get its Cells collection, write a value to A1, and save the workbook as an XLSX file.
// Keywords: Aspose.Cells | C# worksheet name indexer | retrieve Cells collection | write cell value Aspose | save workbook XLSX | Aspose.Cells .NET example | select worksheet by name | Cells API | Excel automation C#
// Common Searches: Aspose.Cells select worksheet using name | How to get Cells object from a specific sheet in C# | C# Aspose.Cells write to cell A1 after selecting sheet | Retrieve worksheet by name indexer Aspose.Cells | Save workbook after modifying cells Aspose
// Developer Intent: Obtain the Cells collection of a named worksheet to perform further read/write operations.
// Use Cases: Insert header text into cell A1 of a dynamically named report sheet. | Loop through all cells of a selected worksheet to apply conditional formatting. | Extract data from a specific sheet for export to a database or API.
// AI Prompts: Generate C# code that selects a worksheet named "Summary" with Aspose.Cells, iterates over its Cells, and sets a yellow background for each cell. | Provide an example that copies the entire Cells collection from a worksheet called "Source" to a new worksheet "Destination" using Aspose.Cells for .NET. | Explain how to safely retrieve a worksheet by name, check for its existence, and handle the case where the sheet is missing in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Shows how to create a workbook, add a sheet called "Report", retrieve that sheet using the name indexer, get its Cells collection, write a value to A1, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Rename the default worksheet
        workbook.Worksheets[0].Name = "DataSheet";

        // Add another worksheet with a specific name
        workbook.Worksheets.Add("Report");

        // Select the worksheet by its name
        Worksheet worksheet = workbook.Worksheets["Report"]; // name indexer

        // Obtain the Cells collection of the selected worksheet
        Cells cells = worksheet.Cells;

        // Example operation on the cells collection
        cells["A1"].PutValue("Hello from Report sheet");

        // Save the workbook
        workbook.Save("SelectWorksheetDemo.xlsx", SaveFormat.Xlsx);
    }
}
