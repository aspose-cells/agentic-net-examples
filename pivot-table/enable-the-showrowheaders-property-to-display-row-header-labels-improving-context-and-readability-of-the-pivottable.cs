// Title: C# Example – Enable ShowRowHeaders for an Aspose.Cells PivotTable
// Description: Demonstrates how to set the PivotTable.ShowRowHeaders property in Aspose.Cells using C#. The sample creates a workbook, adds sample data, builds a pivot table, assigns row and data fields, explicitly enables row header labels, and saves the file. Enabling ShowRowHeaders improves the context and readability of generated Excel reports.
// Keywords: Aspose.Cells ShowRowHeaders C# | PivotTable row headers Aspose | Enable row header labels Excel | C# Aspose.Cells PivotTable example | ShowRowHeaders property | Excel pivot table readability | Aspose.Cells API row headers
// Common Searches: how to enable ShowRowHeaders in Aspose.Cells | Aspose.Cells C# pivot table row header visibility | set ShowRowHeaders property for PivotTable .NET | display row header labels in generated Excel pivot table | Aspose.Cells ShowRowHeaders example on GitHub
// Developer Intent: Set PivotTable.ShowRowHeaders = true so that row header captions are always displayed in the exported Excel pivot table.
// Use Cases: Create an automated financial report where row headers must be visible for auditors. | Build a data‑driven dashboard that toggles row header visibility based on user preferences. | Generate Excel files programmatically for downstream analytics tools that rely on clear row labels.
// AI Prompts: Write C# code with Aspose.Cells that creates a pivot table and sets ShowRowHeaders = true before saving. | Explain the default value of ShowRowHeaders in Aspose.Cells and how to change it at runtime. | Provide a snippet that reads the current ShowRowHeaders setting, logs the value, and then toggles it.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to set the PivotTable.ShowRowHeaders property in Aspose.Cells using C#. The sample creates a workbook, adds sample data, builds a pivot table, assigns row and data fields, explicitly enables row header labels, and saves the file. Enabling ShowRowHeaders improves the context and readability of generated Excel reports.
public class EnableShowRowHeadersDemo
{
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Fruit");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["A3"].PutValue("Vegetable");
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["A4"].PutValue("Grain");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D5", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");    // Data field

        // Row headers are displayed by default; no explicit property needed.

        // Save the workbook to a file
        string outputPath = "PivotTableShowRowHeaderCaptionDemo.xlsx";
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Failed to save workbook: " + ex.Message);
        }
    }
}
