// Title: How to set a PivotTable RowHeaderCaption to a localized string using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pivot table row header caption to Chinese | change RowHeaderCaption of an Excel PivotTable with Aspose.Cells | localize pivot table row header text in .NET using Aspose.Cells | programmatically update PivotTable row header caption in C# | Aspose.Cells example for customizing PivotTable captions for international users
// Tags: Aspose.Cells PivotTable RowHeaderCaption property | C# Excel pivot table caption localization | custom row header caption Aspose.Cells | international Excel pivot table captions .NET | modify pivot table header text programmatically

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads or creates a workbook, obtains or adds a PivotTable on the first worksheet, assigns a Chinese string to the PivotTable's RowHeaderCaption property, and saves the result, demonstrating how to localize pivot table row headers with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load existing workbook or create a sample one if the file is missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = CreateSampleWorkbook();
                workbook.Save(inputPath);
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve or create a PivotTable
            PivotTable pivotTable;
            if (worksheet.PivotTables.Count > 0)
            {
                pivotTable = worksheet.PivotTables[0];
            }
            else
            {
                // Define source data range (A1:C5) for the new PivotTable
                string sourceData = "A1:C5";
                string destinationCell = "E1";
                string pivotName = "PivotTable1";

                // Add a new PivotTable and obtain its index
                int pivotIndex = worksheet.PivotTables.Add(sourceData, destinationCell, pivotName);
                pivotTable = worksheet.PivotTables[pivotIndex];

                // Example: add fields to the PivotTable using PivotField objects
                // Retrieve fields from the source data
                PivotField categoryField = pivotTable.RowFields[0];
                PivotField subCategoryField = pivotTable.ColumnFields[0];
                PivotField amountField = pivotTable.DataFields[0];

                // Add fields to respective collections
                pivotTable.RowFields.Add(categoryField);      // Category
                pivotTable.ColumnFields.Add(subCategoryField); // SubCategory
                pivotTable.DataFields.Add(amountField);       // Amount
            }

            // Set the RowHeaderCaption to a localized string (Chinese example)
            string localizedCaption = "行标题";
            pivotTable.RowHeaderCaption = localizedCaption;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Creates a simple workbook with sample data for PivotTable creation
    private static Workbook CreateSampleWorkbook()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("SubCategory");
        ws.Cells["C1"].PutValue("Amount");

        ws.Cells["A2"].PutValue("Food");
        ws.Cells["B2"].PutValue("Fruit");
        ws.Cells["C2"].PutValue(120);

        ws.Cells["A3"].PutValue("Food");
        ws.Cells["B3"].PutValue("Vegetable");
        ws.Cells["C3"].PutValue(80);

        ws.Cells["A4"].PutValue("Beverage");
        ws.Cells["B4"].PutValue("Tea");
        ws.Cells["C4"].PutValue(50);

        ws.Cells["A5"].PutValue("Beverage");
        ws.Cells["B5"].PutValue("Coffee");
        ws.Cells["C5"].PutValue(70);

        return wb;
    }
}
