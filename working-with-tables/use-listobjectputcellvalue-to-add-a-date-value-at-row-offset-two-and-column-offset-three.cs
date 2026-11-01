// Title: Add a DateTime to a specific cell in an Aspose.Cells ListObject using row and column offsets (C#)
// AI Prompts: Write C# code that creates a workbook, defines a ListObject, and uses ListObject.DataRange.PutValue to insert a DateTime at row offset 2, column offset 3. | Demonstrate how to let the ListObject grow its data range when inserting a value beyond its current bounds using Aspose.Cells. | Provide a complete example that saves the workbook to a file after updating the table cell with a DateTime.
// Common Searches: how to insert a datetime into an Aspose.Cells ListObject table cell at specific offset in C# | Aspose.Cells C# put date value in table using row offset and column offset | how to make a ListObject grow its data range when adding new rows with Aspose.Cells .NET | C# example of ListObject.DataRange.PutValue for date insertion | saving workbook after modifying ListObject cell Aspose.Cells
// Tags: Aspose.Cells ListObject DataRange.PutValue date | C# insert DateTime into Excel table cell | ListObject expansion handling Aspose.Cells | save workbook after ListObject modification .NET | row offset column offset cell update Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// // Creates a workbook, adds a ListObject named 'MyTable' over range A1:D5, inserts a DateTime (2023-12-31) at row offset 2 and column offset 3 within the table using DataRange.PutValue (which expands the table if needed), and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the ListObject (e.g., A1:D5) as a string
            string tableRange = "A1:D5";

            // Add a ListObject (table) to the worksheet with a name
            int listObjectIndex = sheet.ListObjects.Add("MyTable", tableRange, true);
            ListObject listObject = sheet.ListObjects[listObjectIndex];

            // Add a date value at row offset 2 and column offset 3 within the table's data range
            DateTime dateValue = new DateTime(2023, 12, 31);
            // Ensure the data range has enough rows/columns; Aspose.Cells will expand if needed
            listObject.DataRange[2, 3].PutValue(dateValue);

            // Save the workbook
            string outputPath = "Output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
