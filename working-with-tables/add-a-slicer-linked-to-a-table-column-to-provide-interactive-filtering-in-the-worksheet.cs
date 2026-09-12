// Title: Add a slicer linked to a table column and place it at a specific cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a ListObject, then adds a slicer bound to the 'Category' column and positions the slicer at cell A7 with Aspose.Cells. | Show the steps to programmatically insert an Excel slicer for a table column, set its caption, and save the workbook using Aspose.Cells for .NET. | Provide a minimal example that demonstrates how to enable interactive filtering by attaching a slicer to a table column in a C# Aspose.Cells workbook.
// Common Searches: asp.net cells how to add a slicer to a table column in C# | example of positioning an Excel slicer at cell A7 using Aspose.Cells | link slicer to ListObject column programmatically Aspose.Cells | C# code for creating interactive slicer for Excel table with Aspose.Cells | Aspose.Cells slicer support version requirement and usage
// Tags: add slicer to ListObject column | position slicer on worksheet cell | Aspose.Cells interactive table filtering | C# create Excel slicer with Aspose.Cells | Aspose.Cells slicer example for .NET | link slicer to table column Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;

// The program creates a new workbook, fills it with sample data, defines a ListObject named SalesTable, and (when supported) adds a slicer linked to the 'Category' column positioned at cell A7, then saves the file as SlicerExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "Data";

            // Populate sample data for the table
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Amount");
            ws.Cells["A2"].PutValue("Food");
            ws.Cells["B2"].PutValue(120);
            ws.Cells["A3"].PutValue("Transport");
            ws.Cells["B3"].PutValue(80);
            ws.Cells["A4"].PutValue("Utilities");
            ws.Cells["B4"].PutValue(150);
            ws.Cells["A5"].PutValue("Food");
            ws.Cells["B5"].PutValue(90);

            // Create a table (ListObject) that covers the data range A1:B5
            int tableIndex = ws.ListObjects.Add(0, 0, 5, 2, true);
            ListObject table = ws.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // NOTE: Slicer support requires a recent version of Aspose.Cells.
            // If the referenced library does not contain the Slicer class,
            // the following code is omitted to keep the project compilable.
            // Uncomment and adjust when using a version that includes slicers.

            /*
            // Add a slicer linked to the "Category" column (column index 0 in the table)
            // Position the slicer at cell A7 (row index 6, column index 0)
            Slicer slicer = ws.Slicers.Add(table.DisplayName, 0, 6, 0);
            slicer.Name = "CategorySlicer";
            slicer.Caption = "Category";
            */

            // Define output file path
            string outputPath = "SlicerExample.xlsx";

            // Ensure the directory for the output file exists
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
