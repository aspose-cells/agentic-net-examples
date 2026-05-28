using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class PivotItemPositionInSameParentNodeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputFile = "PivotTest3.xlsx";
            const string outputFile = "PivotItemPositionInSameParentNode_Output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputFile);

                // Access the worksheet that holds the data (ensure it exists)
                Worksheet dataSheet = workbook.Worksheets["New Hardware - Yearly"];
                if (dataSheet == null)
                {
                    Console.WriteLine("Worksheet \"New Hardware - Yearly\" not found.");
                    return;
                }

                // Access the worksheet that contains the pivot table
                Worksheet pivotSheet = workbook.Worksheets["PivotTable"];
                if (pivotSheet == null)
                {
                    Console.WriteLine("Worksheet \"PivotTable\" not found.");
                    return;
                }

                // Ensure at least one pivot table exists on the sheet
                if (pivotSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the \"PivotTable\" sheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = pivotSheet.PivotTables[0];

                // Retrieve the row field named "Item"
                PivotField itemField = pivotTable.RowFields["Item"];
                if (itemField == null)
                {
                    Console.WriteLine("Row field \"Item\" not found in the pivot table.");
                    return;
                }

                // Iterate through all pivot items in the field and display their positions
                foreach (PivotItem item in itemField.PivotItems)
                {
                    int position = item.PositionInSameParentNode;
                    Console.WriteLine($"Item: {item.Name}, PositionInSameParentNode: {position}");
                }

                // Save the workbook (if any changes were made)
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}