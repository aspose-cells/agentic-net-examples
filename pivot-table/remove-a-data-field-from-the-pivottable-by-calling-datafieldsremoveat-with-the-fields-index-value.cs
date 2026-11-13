// Title: Remove a data field from an Aspose.Cells PivotTable in C# using DataFields.RemoveAt
// Common Searches: C# code to delete a specific data field from an Aspose.Cells PivotTable | How to use DataFields.RemoveAt to remove pivot table data fields in .NET | Aspose.Cells remove pivot table data field by index example | Programmatically delete a pivot table data field in a .xlsx workbook using C# | Aspose.Cells PivotTable DataFields.RemoveAt usage guide for developers
// Tags: Aspose.Cells PivotTable data field removal | C# DataFields.RemoveAt method | modify Excel PivotTable data fields Aspose.Cells | .NET remove pivot table data field by index | Aspose.Cells workbook pivot table manipulation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

namespace PivotTableExample
{
    // The example loads an existing workbook, accesses the first worksheet, retrieves a PivotTable named "PivotTable1", validates a zero‑based data field index, removes the data field at that index with DataFields.RemoveAt, and saves the updated workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook that contains a PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Access the worksheet that holds the PivotTable (first worksheet in this example)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the PivotTable by its name (replace "PivotTable1" with your actual name)
                PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
                if (pivotTable == null)
                {
                    Console.WriteLine("Error: PivotTable \"PivotTable1\" was not found on the worksheet.");
                    return;
                }

                // Index of the data field to remove (0‑based). Adjust as needed.
                int dataFieldIndex = 0;

                // Ensure the index is within the range of existing data fields
                if (dataFieldIndex < 0 || dataFieldIndex >= pivotTable.DataFields.Count)
                {
                    Console.WriteLine("Error: Specified data field index is out of range.");
                    return;
                }

                // Remove the data field at the specified index
                pivotTable.DataFields.RemoveAt(dataFieldIndex);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
