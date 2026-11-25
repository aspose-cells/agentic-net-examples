// Title: How to use Label and LabelPosition attributes with Aspose.Cells smart markers to add section headings before grouped rows in C#
// Common Searches: Aspose.Cells C# set Label attribute on smart marker to create a heading for a grouped row range | using LabelPosition property in Aspose.Cells smart markers to position section titles above data rows | C# example of adding group labels with smart markers before collapsed rows in Excel using Aspose.Cells | how to display custom headings for grouped rows with smart markers in Aspose.Cells for .NET | Aspose.Cells smart markers labelposition usage for Excel report section headers
// Tags: smart markers label attribute Aspose.Cells | labelposition property Aspose.Cells C# | group rows with smart markers Excel | section headings for collapsed rows Aspose.Cells | C# generate Excel report with smart marker headings

using Aspose.Cells;
using System;

// The example creates a workbook, populates it with sample data, groups specific rows, and shows how to assign the Label and LabelPosition attributes on smart markers so that a heading appears before each grouped row block, then saves the file as GroupedData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (header + rows)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("Item 4");
            sheet.Cells["B5"].PutValue(40);

            // Group rows 2‑3 (zero‑based indices 1‑2) and collapse the group
            sheet.Cells.GroupRows(1, 2, true);

            // Group rows 4‑5 (zero‑based indices 3‑4) and collapse the group
            sheet.Cells.GroupRows(3, 4, true);

            // Save the workbook to a file
            string outputPath = "GroupedData.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
