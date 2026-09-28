// Title: C# utility method to delete a named pivot table from an Excel workbook using Aspose.Cells
// AI Prompts: Write a C# function that receives an input Excel file path, a pivot table name, and an output path, then removes the matching pivot table from every worksheet with Aspose.Cells and saves the workbook. | Update the deletion routine to return a boolean value indicating whether the specified pivot table was found and successfully removed. | Create an overload that accepts a list of pivot table names and deletes all matching tables from the workbook in one pass.
// Common Searches: how to remove a specific pivot table by its name using Aspose.Cells in C# | Aspose.Cells C# delete pivot table from all worksheets example | programmatically delete pivot table from Excel file with Aspose.Cells .NET | C# code to find pivot table index and call RemoveAt in Aspose.Cells | remove multiple pivot tables by name Aspose.Cells C#
// Tags: Aspose.Cells PivotTableCollection.RemoveAt example | delete named pivot table C# Aspose.Cells | remove pivot table by index Excel .NET | Excel workbook pivot table deletion using Aspose.Cells | C# utility to purge specific pivot tables

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Provides a C# method that locates a pivot table by its name across all worksheets in an Excel workbook, removes it using PivotTableCollection.RemoveAt, and saves the modified file to a specified output path.
public static class PivotTableUtility
{
    /// <param name="inputFilePath">Path to the source Excel file.</param>
    /// <param name="pivotTableName">Name of the pivot table to delete.</param>
    /// <param name="outputFilePath">Path where the modified workbook will be saved.</param>
    public static void DeletePivotTableByName(string inputFilePath, string pivotTableName, string outputFilePath)
    {
        try
        {
            // Verify that the input file exists.
            if (!File.Exists(inputFilePath))
                throw new FileNotFoundException($"Input file not found: {inputFilePath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputFilePath);

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet.
                PivotTableCollection pivotTables = worksheet.PivotTables;

                // Search for the pivot table by name.
                for (int i = 0; i < pivotTables.Count; i++)
                {
                    PivotTable pt = pivotTables[i];
                    if (string.Equals(pt.Name, pivotTableName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the pivot table at the found index.
                        pivotTables.RemoveAt(i);
                        // Exit the loop after removal; pivot table names are unique per worksheet.
                        break;
                    }
                }
            }

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook to the output path.
            workbook.Save(outputFilePath);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error deleting pivot table: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    /// <summary>
    /// Entry point for the console application.
    /// Usage: <executable> <inputFilePath> <pivotTableName> <outputFilePath>
    /// </summary>
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.WriteLine("Usage: <executable> <inputFilePath> <pivotTableName> <outputFilePath>");
            return;
        }

        string inputFilePath = args[0];
        string pivotTableName = args[1];
        string outputFilePath = args[2];

        try
        {
            PivotTableUtility.DeletePivotTableByName(inputFilePath, pivotTableName, outputFilePath);
            Console.WriteLine($"Pivot table \"{pivotTableName}\" deleted successfully. Output saved to \"{outputFilePath}\".");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
