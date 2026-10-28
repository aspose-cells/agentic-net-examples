// Title: Create a nested list with two-level indentation in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to add two-level indentation to Excel cells with Aspose.Cells in C# | populate hierarchical list with subitems in Excel using Aspose.Cells .NET | write indented rows in an Excel workbook programmatically with Aspose.Cells C# | generate nested list with leading spaces in cells using Aspose.Cells for .NET | save Excel file containing indented list using Aspose.Cells C# example
// Tags: add leading spaces to cell values Aspose.Cells C# | populate hierarchical list Excel Aspose.Cells .NET | write indented rows workbook Aspose.Cells | save workbook with nested list Aspose.Cells | Excel cell indentation using Aspose.Cells API

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, fills column A with items and sub‑items using leading spaces to represent two levels of indentation, and saves the file as NestedListWithIndentation.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data representing a nested list
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                // Add indentation manually by prefixing spaces (two levels)
                sheet.Cells["A4"].PutValue("  Subitem 1.1");
                sheet.Cells["A5"].PutValue("  Subitem 1.2");
                sheet.Cells["A6"].PutValue("  Subitem 2.1");
                sheet.Cells["A7"].PutValue("  Subitem 2.2");

                // Save the workbook
                string outputPath = "NestedListWithIndentation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
