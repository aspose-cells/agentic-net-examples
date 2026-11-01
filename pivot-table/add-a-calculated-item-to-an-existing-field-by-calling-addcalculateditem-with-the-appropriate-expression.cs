// Title: How to add a calculated item to a pivot table row field using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add calculated item to pivot row field example | Create a calculated item in Excel pivot table with AddCalculatedItem in .NET | C# code to insert a calculated item into an existing pivot table using Aspose.Cells | Define a calculated item expression for a pivot field in Aspose.Cells workbook | Add calculated item HighValue to Category field in Excel pivot via Aspose.Cells
// Tags: AddCalculatedItem Aspose.Cells C# | pivot table calculated item .NET | Excel pivot row field calculated item | Aspose.Cells workbook pivot manipulation | C# expression for calculated pivot item

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The program loads an existing workbook, verifies a pivot table is present, retrieves the 'Category' row field, adds a calculated item named 'HighValue' with the expression ='Category'="Premium", and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table (adjust index if needed)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Get the row field named "Category"
            PivotField field = pivotTable.RowFields["Category"];
            if (field == null)
            {
                Console.WriteLine("Pivot field \"Category\" not found.");
                return;
            }

            // Add a calculated item named "HighValue" with the desired expression
            // The expression must be a valid Excel formula for a calculated item.
            field.AddCalculatedItem("HighValue", "='Category'=\"Premium\"");

            // Save the workbook with the new calculated item
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
