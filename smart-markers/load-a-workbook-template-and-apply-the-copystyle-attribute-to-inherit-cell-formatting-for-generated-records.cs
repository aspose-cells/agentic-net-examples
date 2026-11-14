// Title: Insert rows with inherited formatting into an Excel template using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert row while keeping original cell style | How to use CopyStyle when adding rows to an Excel template with Aspose.Cells | Load an Excel template and populate data rows preserving formatting in .NET | Example of InsertRows with CopyStyle = true in Aspose.Cells C# | Add records to a workbook template without losing cell formatting using Aspose.Cells
// Tags: insert rows copystyle Aspose.Cells | load workbook template C# Aspose.Cells | preserve cell formatting during row insertion | populate Excel template with data Aspose.Cells | copy style attribute InsertRows Aspose.Cells

using Aspose.Cells;
using System.Collections.Generic;

// The example loads an Excel template (Template.xlsx), iterates over a list of records, inserts new rows with the CopyStyle flag set to true so the new rows inherit the formatting of the previous row, writes the record values into the inserted cells, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook template
        var workbook = new Workbook("Template.xlsx");
        var worksheet = workbook.Worksheets[0];

        // Sample data to be added
        var records = new List<string[]>
        {
            new[] { "John", "Doe", "30" },
            new[] { "Jane", "Smith", "25" }
        };

        // Assume the template has a header row at index 0 and a sample data row at index 1
        int insertRowIndex = 1; // Row where new records start

        foreach (var record in records)
        {
            // Insert a new row and copy the style from the row above (CopyStyle = true)
            worksheet.Cells.InsertRows(insertRowIndex, 1, true);

            // Fill the inserted row with values
            for (int col = 0; col < record.Length; col++)
            {
                worksheet.Cells[insertRowIndex, col].PutValue(record[col]);
            }

            // Move to the next row position for the following record
            insertRowIndex++;
        }

        // Save the workbook with the generated records
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}
