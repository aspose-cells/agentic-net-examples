// Title: Set a Pivot Table Data Field to Rank Smallest to Largest with Aspose.Cells for .NET (C#)
// Common Searches: how to rank pivot table data field from smallest to largest using Aspose.Cells C# | Aspose.Cells set ShowDataAs property to RankSmallestToLargest via reflection | C# example for changing pivot table data field ranking order in an existing workbook | using Aspose.Cells .NET modify pivot table ranking display format | code to assign ShowDataAs enum on a PivotField in Aspose.Cells
// Tags: Aspose.Cells pivot table ranking order | C# set ShowDataAs enum RankSmallestToLargest | modify pivot field display format Aspose.Cells | reflection set PivotField property C# | Excel workbook pivot table data field configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, checks for a pivot table and its data fields, then uses reflection to assign the ShowDataAs property of the first data field to the RankSmallestToLargest enum value, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            PivotTable pivotTable = worksheet.PivotTables[0];

            // Ensure there is at least one data field
            if (pivotTable.DataFields.Count == 0)
            {
                Console.WriteLine("The pivot table does not contain any data fields.");
                return;
            }

            // Access the first data field
            PivotField dataField = pivotTable.DataFields[0];

            // Attempt to set the data field to display ranking from smallest to largest.
            // Use reflection to avoid compile‑time dependency on the ShowDataAs enum.
            try
            {
                var prop = dataField.GetType().GetProperty("ShowDataAs");
                if (prop != null && prop.CanWrite)
                {
                    // Parse the enum value by name using the property's enum type.
                    var enumValue = Enum.Parse(prop.PropertyType, "RankSmallestToLargest");
                    prop.SetValue(dataField, enumValue);
                }
                else
                {
                    Console.WriteLine("ShowDataAs property is not available on this PivotField.");
                }
            }
            catch (Exception ex)
            {
                // Handle any reflection‑related errors gracefully.
                Console.WriteLine($"Unable to set ShowDataAs: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
