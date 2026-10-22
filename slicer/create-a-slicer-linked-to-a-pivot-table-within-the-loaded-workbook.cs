// Title: Insert a slicer linked to a pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer for a specified pivot field, positions it at a given cell, and saves the workbook. | Show how to modify slicer properties such as Name, Caption, Width, and Height after the slicer is created with Aspose.Cells.
// Common Searches: how to programmatically add a slicer to a pivot table with Aspose.Cells C# | Aspose.Cells example for linking slicer to pivot field and setting its size | C# code to place an Excel slicer at cell A1 using Aspose.Cells | customizing slicer caption and name in Aspose.Cells .NET
// Tags: Aspose.Cells add slicer to pivot table | C# set slicer properties Aspose.Cells | Aspose.Cells slicer placement cell A1 | Excel slicer customization Aspose.Cells .NET | save workbook after inserting slicer Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example loads an existing Excel file, retrieves the first pivot table, adds a slicer linked to the "Category" field at cell A1, customizes its name, caption, width, and height, and then saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table is the first one on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Field name to be used for the slicer (must match a pivot field)
            string slicerFieldName = "Category"; // replace with your actual field name

            // Add a slicer linked to the pivot table at cell A1
            int slicerIndex = worksheet.Slicers.Add(pivotTable, "A1", slicerFieldName);
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Optional: customize slicer appearance
            slicer.Name = "CategorySlicer";
            slicer.Caption = "Select Category";
            slicer.Width = 150;
            slicer.Height = 200;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
