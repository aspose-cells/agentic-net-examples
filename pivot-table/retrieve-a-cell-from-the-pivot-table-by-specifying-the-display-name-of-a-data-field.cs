// Title: Retrieve a PivotTable cell by its data field display name with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get value of a pivot table data field by display name | C# retrieve first data row cell for a specific pivot table field using Aspose.Cells | How to locate a pivot table data field index and read its cell value in .NET | Using reflection to access PivotTable.DataRange in Aspose.Cells C# example | Read 'Sum of Amount' value from the first row of a pivot table with Aspose.Cells
// Tags: Aspose.Cells PivotTable data field index retrieval | Aspose.Cells read pivot table cell value | C# get PivotTable DataRange via reflection | retrieve pivot table cell by display name | Aspose.Cells first data row value extraction

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, accesses the first worksheet and its first PivotTable, finds the index of a data field by its display name, obtains the PivotTable's DataRange through reflection, calculates the cell coordinates for the first data row of that field, reads the cell value, and prints it to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table (or use a specific name/index)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Display name of the data field to retrieve
            string dataFieldDisplayName = "Sum of Amount";

            // Locate the data field index within the pivot table's DataFields collection
            int dataFieldIndex = -1;
            for (int i = 0; i < pivotTable.DataFields.Count; i++)
            {
                PivotField df = pivotTable.DataFields[i];
                if (df.Name == dataFieldDisplayName)
                {
                    dataFieldIndex = i;
                    break;
                }
            }

            if (dataFieldIndex == -1)
            {
                Console.WriteLine($"Data field '{dataFieldDisplayName}' not found.");
                return;
            }

            // Obtain the data range via reflection (covers versions where DataRange may be missing)
            PropertyInfo dataRangeProp = typeof(PivotTable).GetProperty("DataRange", BindingFlags.Public | BindingFlags.Instance);
            if (dataRangeProp == null)
            {
                Console.WriteLine("The PivotTable.DataRange property is not available in this Aspose.Cells version.");
                return;
            }

            // Use fully qualified Aspose.Cells.Range to avoid ambiguity with System.Range
            Aspose.Cells.Range dataRange = dataRangeProp.GetValue(pivotTable) as Aspose.Cells.Range;
            if (dataRange == null)
            {
                Console.WriteLine("Unable to retrieve the data range of the pivot table.");
                return;
            }

            // Calculate the cell coordinates for the first data row of the selected field
            int targetRow = dataRange.FirstRow + 1;               // skip header row
            int targetColumn = dataRange.FirstColumn + dataFieldIndex;

            // Retrieve the cell value
            Cell targetCell = worksheet.Cells[targetRow, targetColumn];
            Console.WriteLine($"Value of '{dataFieldDisplayName}' in the first data row: {targetCell.Value}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
