// Title: How to enable repeating row item labels in an Aspose.Cells PivotTable using C#
// AI Prompts: Generate C# code that creates a workbook, adds sample data, builds a pivot table, and sets PivotTable.RowFields[0].IsRepeatItemLabels = true with Aspose.Cells. | Show the steps to modify an existing Aspose.Cells pivot table in .NET to turn on repeat item labels for every row field. | Provide a complete example that saves the workbook after enabling repeat item labels for pivot table rows using Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# repeat item labels in pivot table rows example | set IsRepeatItemLabels true for pivot table row fields Aspose.Cells .NET | how to make pivot table row labels repeat in generated Excel file using Aspose.Cells | C# code to enable repeating row item labels in Aspose.Cells pivot table
// Tags: Aspose.Cells pivot table repeat item labels | C# set IsRepeatItemLabels property | enable row field label repetition Aspose.Cells | create pivot table with repeating row labels .NET | Aspose.Cells workbook pivot table configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates creating a workbook, adding sample data, building a pivot table, enabling repeating row item labels by setting IsRepeatItemLabels to true, refreshing the pivot, and saving the file as an .xlsx using Aspose.Cells for C#.
public class EnableRepeatItemLabelsDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // Add a pivot table to the worksheet
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add a row field and a data field to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Enable repeating item labels for the first row field
            pivotTable.RowFields[0].IsRepeatItemLabels = true;

            // Refresh the pivot table data
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("EnableRepeatItemLabelsDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        EnableRepeatItemLabelsDemo.Run();
    }
}
