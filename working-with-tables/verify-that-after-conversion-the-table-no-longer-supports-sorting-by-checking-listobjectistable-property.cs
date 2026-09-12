// Title: Determine whether a ListObject remains a table after workbook conversion by reading its IsTable property with Aspose.Cells for .NET
// AI Prompts: Using Aspose.Cells in C#, open a converted Excel file, locate the first ListObject on the first worksheet, and print the value of its IsTable property to confirm if the object still behaves as a table. | Write C# code that loads a workbook, checks for ListObjects, and outputs whether each ListObject's IsTable flag is true, indicating that sorting is still supported. | Create a method that accepts a file path, loads the workbook with Aspose.Cells, and returns a boolean indicating if the first table can be sorted based on the ListObject.IsTable property.
// Common Searches: Aspose.Cells check ListObject.IsTable after converting Excel file to .xlsx | C# verify if Excel table still supports sorting using Aspose.Cells IsTable property | How to determine if a ListObject is a table after workbook conversion in Aspose.Cells .NET | Read ListObject.IsTable flag to confirm table status in converted Excel workbook | Aspose.Cells ListObject sorting support detection C#
// Tags: Aspose.Cells ListObject IsTable verification | C# detect table status after workbook conversion | check Excel table sorting capability Aspose.Cells | listobject property inspection .NET | validate table existence after file conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

// The example loads a workbook that has been converted to .xlsx, accesses the first worksheet, ensures a ListObject exists, reads its IsTable boolean property, and writes the result to the console to confirm whether the object still functions as a sortable table.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook that contains a table (ListObject)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one ListObject (table)
            if (sheet.ListObjects.Count > 0)
            {
                // Retrieve the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Example property check: whether the table header row is displayed
                bool showHeader = table.ShowHeaderRow;

                // Output the result
                Console.WriteLine($"ListObject.ShowHeaderRow: {showHeader}");
            }
            else
            {
                Console.WriteLine("No ListObjects (tables) found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
