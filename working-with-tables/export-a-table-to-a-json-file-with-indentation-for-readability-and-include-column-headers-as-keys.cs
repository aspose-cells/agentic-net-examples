// Title: Export an Excel worksheet table to indented JSON with column headers as keys using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet to pretty printed JSON with column names as keys | Save Excel table as formatted JSON file using Aspose.Cells .NET | Generate indented JSON from Excel data with column headers using Aspose.Cells in C# | C# code to convert an Excel range to readable JSON with indentation via Aspose.Cells | Aspose.Cells JsonSaveOptions example for readable JSON output
// Tags: Aspose.Cells JsonSaveOptions pretty-printed output | export worksheet to JSON C# | Excel to JSON conversion with column headers Aspose.Cells | indented JSON generation from Excel Aspose.Cells | C# Aspose.Cells save as JSON file

using System;
using Aspose.Cells;

// The sample creates a workbook, adds a worksheet named "Data" with column headers and sample rows, applies JsonSaveOptions, and saves the worksheet content as an indented JSON file (TableData.json) where each column header becomes a JSON key.
class ExportTableToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate column headers
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Age");
            sheet.Cells["D1"].PutValue("Country");

            // Populate sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["D2"].PutValue("USA");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Emma");
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["D3"].PutValue("Canada");

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Liu");
            sheet.Cells["C4"].PutValue(28);
            sheet.Cells["D4"].PutValue("China");

            // Use default JSON save options
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Save the workbook data to a JSON file
            workbook.Save("TableData.json", jsonOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
