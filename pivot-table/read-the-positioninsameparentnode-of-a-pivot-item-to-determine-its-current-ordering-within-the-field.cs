// Title: Read the PositionInSameParentNode of each PivotItem in a row field using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens a workbook, accesses the first pivot table, iterates over the row field's PivotItems, and prints each item's PositionInSameParentNode. | Show how to retrieve the ordering index of pivot items in a row field with Aspose.Cells and output the item names together with their positions. | Provide an example that reads PivotItem.PositionInSameParentNode, logs the results, and then saves the workbook. | Explain how to safely handle missing pivot tables or row fields when extracting PivotItem positions in Aspose.Cells.
// Common Searches: Aspose.Cells C# get PositionInSameParentNode for pivot items in a row field | How to determine the order of PivotItems in a pivot table using Aspose.Cells .NET | Iterate over row field items and read their position index with Aspose.Cells | C# example for reading pivot item ordering from an existing Excel file | Aspose.Cells read pivot item position without modifying the workbook
// Tags: aspnet read pivotitem positioninsameparentnode | aspose.cells iterate row field items | pivot table item ordering aspnet | c# retrieve pivot item index aspose.cells | excel workbook pivot item position reading

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing Excel workbook, accesses the first worksheet and its first pivot table, selects the first row field, iterates through all PivotItems of that field, reads each item's PositionInSameParentNode to determine its ordering within the parent node, prints the item name and position, and finally saves the workbook (optional).
public class PivotItemPositionReader
{
    public static void Run()
    {
        const string inputFile = "PivotTest3.xlsx";
        const string outputFile = "PivotItemPositionReader_Output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first pivot table on the worksheet
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure there is at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Error: Pivot table contains no row fields.");
                return;
            }

            // Choose a row field to examine (here we use the first row field)
            PivotField rowField = pivotTable.RowFields[0];

            // Iterate through all pivot items of the selected row field
            foreach (PivotItem item in rowField.PivotItems)
            {
                // Read the PositionInSameParentNode property which indicates the item's
                // current ordering within its parent node
                int positionInSameParent = item.PositionInSameParentNode;

                // Output the item name and its position
                Console.WriteLine($"Item Name: {item.Name}, PositionInSameParentNode: {positionInSameParent}");
            }

            // Save the workbook (optional, as we only read data)
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        PivotItemPositionReader.Run();
    }
}
