// Title: How to delete a comment from an Excel table (ListObject) using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens a workbook, locates a ListObject by name, clears its Comment property, and saves the file. | Show an example that validates the existence of the input file, worksheet, and table before removing the table comment in Aspose.Cells. | Provide a snippet that demonstrates error handling for missing files or invalid table names while clearing a ListObject comment.
// Common Searches: C# Aspose.Cells clear comment on a ListObject table | example code to remove Excel table comment using Aspose.Cells .NET | how to delete metadata comment from an Excel table programmatically | Aspose.Cells delete table comment after loading workbook
// Tags: Aspose.Cells clear ListObject comment | C# remove Excel table comment | Aspose.Cells delete table comment | programmatic Excel table metadata cleanup .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// This example loads an existing Excel workbook with Aspose.Cells, accesses a specified worksheet and ListObject (table) by name, clears the table's Comment property if present, and saves the workbook to a new file. It includes robust error handling for missing files, worksheets, or tables.
class Program
{
    static void Main()
    {
        // Define file paths and identifiers (replace with actual values)
        string inputFilePath = "{InputFilePath}";
        string outputFilePath = "{OutputFilePath}";
        string sheetName = "{SheetName}";
        string tableName = "{TableName}";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException("Input workbook not found.", inputFilePath);

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFilePath);

            // Access the specified worksheet
            Worksheet sheet = workbook.Worksheets[sheetName];
            if (sheet == null)
                throw new ArgumentException($"Worksheet \"{sheetName}\" does not exist.");

            // Retrieve the table (ListObject) by its name
            ListObject table = sheet.ListObjects[tableName];
            if (table == null)
                throw new ArgumentException($"Table \"{tableName}\" does not exist in worksheet \"{sheetName}\".");

            // Delete the comment attached to the table, if it exists
            if (!string.IsNullOrEmpty(table.Comment))
                table.Comment = string.Empty; // clears the comment

            // Save the workbook after removing the comment
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
