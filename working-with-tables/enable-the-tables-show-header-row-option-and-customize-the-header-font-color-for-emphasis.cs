// Title: How to display a table header row and apply a red font color to the header using Aspose.Cells for .NET
// AI Prompts: Create an Excel workbook, add a ListObject table, enable its ShowHeaderRow property, and set the header row font color to red with Aspose.Cells in C#. | Write C# code that defines a red‑font style, applies it to the first row of a ListObject table, and saves the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# show table header row and change header text color | set ListObject ShowHeaderRow true and style header text red | apply custom font color to Excel table header using Aspose.Cells .NET | how to style table header row in Aspose.Cells workbook | C# Aspose.Cells change header row text color to red
// Tags: Aspose.Cells ListObject header visibility | Aspose.Cells header font color styling | C# style Excel table header Aspose.Cells | Apply style to table header Aspose.Cells | Excel table header customization Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.Drawing;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, adds a ListObject table with a visible header row, defines a style with a red font, applies this style to the header range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including header row)
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Define the range for the table (including header)
            int firstRow = 0;          // zero‑based index for row 1
            int firstColumn = 0;       // zero‑based index for column A
            int totalRows = 3;         // header + 2 data rows
            int totalColumns = 2;      // two columns: Name and Age

            // Add a ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Ensure the header row is displayed
            table.ShowHeaderRow = true;

            // Create a style for the header row and set the font color
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Color = Color.Red; // Emphasize header with red font

            // Apply the style to the header row of the table
            AsposeRange headerRange = sheet.Cells.CreateRange(firstRow, firstColumn, 1, totalColumns);
            headerRange.SetStyle(headerStyle);

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
