// Title: Load an Excel workbook from a file path and list worksheet names – Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a Workbook object from a local Excel file, retrieve the Workbook.Worksheets collection, iterate through each Worksheet to output its name, and properly dispose the workbook to free resources.
// Keywords: Aspose.Cells load workbook C# | open Excel file by path .NET | access worksheets collection Aspose.Cells | enumerate worksheet names C# | dispose Workbook Aspose.Cells | C# Excel automation | global .NET developers | US .NET Excel processing
// Common Searches: How to open an existing Excel file with Aspose.Cells in C# | Aspose.Cells get list of sheet names from workbook | C# code to iterate worksheets after loading workbook | Best way to release Aspose.Cells Workbook resources
// Developer Intent: Open a local Excel file, read its worksheets collection, and loop through the sheets to obtain their names.
// Use Cases: Populate a dropdown with all sheet names after a user uploads an Excel file. | Verify required worksheets (e.g., "Data" and "Summary") exist before data extraction. | Create an audit log of worksheet names for compliance reporting.
// AI Prompts: Write C# code that loads an Excel workbook from a given file path, accesses the Worksheets collection, and prints each worksheet name using Aspose.Cells. | Explain memory‑management best practices for disposing Aspose.Cells Workbook objects in a high‑throughput .NET service. | Generate a unit test that confirms the workbook loads correctly and the expected worksheet names are returned.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Demonstrates how to create a Workbook object from a local Excel file, retrieve the Workbook.Worksheets collection, iterate through each Worksheet to output its name, and properly dispose the workbook to free resources.
    class Program
    {
        static void Main()
        {
            // Path to the existing Excel file
            string filePath = "input.xlsx";

            // Load the workbook from the specified file path using the string constructor
            Workbook workbook = new Workbook(filePath);

            // Access the worksheets collection
            WorksheetCollection worksheets = workbook.Worksheets;

            // Example: iterate through all worksheets and print their names
            for (int i = 0; i < worksheets.Count; i++)
            {
                Worksheet sheet = worksheets[i];
                Console.WriteLine($"Worksheet {i}: {sheet.Name}");
            }

            // Optional: clean up resources
            workbook.Dispose();
        }
    }
}
