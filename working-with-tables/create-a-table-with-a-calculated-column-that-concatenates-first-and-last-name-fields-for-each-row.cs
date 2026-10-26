// Title: Generate an Excel table with a FullName calculated column that concatenates FirstName and LastName using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add calculated column to ListObject that joins first and last name | how to create a concatenated FullName column in an Excel table with Aspose.Cells .NET | C# set cell formula dynamically for each row in an Aspose.Cells table | populate Excel ListObject with sample data and a derived column using Aspose.Cells | Aspose.Cells create table and assign formula to cells in C#
// Tags: Aspose.Cells derived column formula | C# ListObject creation with concatenated field | Excel table full name generation Aspose.Cells | dynamic cell formula assignment in .NET | save workbook as xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds FirstName, LastName, and FullName headers, fills sample names, assigns a formula to each FullName cell that joins the corresponding first and last names with a space, builds a ListObject (Excel table) over the range, and saves the file as PeopleTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set headers
            sheet.Cells["A1"].PutValue("FirstName");
            sheet.Cells["B1"].PutValue("LastName");
            sheet.Cells["C1"].PutValue("FullName");

            // Sample data
            string[] firstNames = { "John", "Jane", "Bob", "Alice", "Mike" };
            string[] lastNames = { "Doe", "Smith", "Brown", "Johnson", "Davis" };

            // Populate FirstName and LastName columns
            for (int i = 0; i < firstNames.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(firstNames[i]); // Column A
                sheet.Cells[i + 1, 1].PutValue(lastNames[i]); // Column B
            }

            // Add calculated column that concatenates FirstName and LastName
            for (int row = 1; row <= firstNames.Length; row++)
            {
                // Formula: =A2 & " " & B2 (adjusted for each row)
                sheet.Cells[row, 2].Formula = $"=A{row + 1} & \" \" & B{row + 1}";
            }

            // Define the range for the table (including header row)
            int totalRows = firstNames.Length + 1; // header + data rows
            int totalCols = 3; // FirstName, LastName, FullName

            // Create a table (ListObject) over the defined range
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows - 1, totalCols - 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "PeopleTable";
            table.ShowHeaderRow = true;
            // ShowTotalRow defaults to false; no need to set it explicitly

            // Determine output path and ensure directory exists
            string outputPath = "PeopleTable.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
