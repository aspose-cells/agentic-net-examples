using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RetrievePivotItemPosition
{
    static void Main()
    {
        const string inputFile = "PivotData.xlsx";
        const string outputFile = "PivotData_Output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputFile)}");
                return;
            }

            // Load the workbook containing the pivot table
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first pivot table on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Obtain the first row field of the pivot table
            PivotField rowField = pivotTable.RowFields[0];

            // Retrieve the first pivot item in the collection
            PivotItem pivotItem = rowField.PivotItems[0];

            // Read the absolute Position property of the pivot item
            int absolutePosition = pivotItem.Position;

            Console.WriteLine($"Pivot Item Name: {pivotItem.Name}");
            Console.WriteLine($"Absolute Position: {absolutePosition}");

            // Save the workbook (optional, if any changes were made)
            workbook.Save(outputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}