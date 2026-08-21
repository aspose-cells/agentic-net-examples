// Title: Import a sliced array segment into Excel with Aspose.Cells C# ImportArray and range operator
// Description: Demonstrates how to use the C# range operator (e.g., orders[1..4]) to extract a subset of an array, convert it to a primitive array, and import the values vertically into a worksheet using Aspose.Cells Worksheet.Cells.ImportArray.
// Keywords: Aspose.Cells | ImportArray | C# range operator | array slice | subset import | Excel export C# | smart markers alternative | worksheet.Cells.ImportArray | C# 8.0 | Excel data import
// Common Searches: Aspose.Cells import sliced array | C# range operator ImportArray example | How to import part of an array into Excel with Aspose | ImportArray with C# 8 range syntax | Excel export only selected collection items Aspose.Cells
// Developer Intent: Import a specific portion of an in‑memory array into an Excel worksheet using Aspose.Cells.
// Use Cases: Export only the middle rows of a large dataset to keep reports concise. | Create a dynamic summary sheet that shows a sliding window of values based on user‑selected indices. | Generate Excel files where the data range is calculated at runtime, avoiding the need to load the full collection.
// AI Prompts: Modify the code to import orders[2..5] into column C starting at C2. | Show how to achieve the same sliced import using Aspose.Cells smart markers instead of ImportArray. | Add validation that the requested slice does not exceed the source array bounds and return a friendly error message.

using Aspose.Cells;
using System;
using System.Linq;

// Demonstrates how to use the C# range operator (e.g., orders[1..4]) to extract a subset of an array, convert it to a primitive array, and import the values vertically into a worksheet using Aspose.Cells Worksheet.Cells.ImportArray.
public class SubsetImportDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data source: an array of Order objects
            Order[] orders = new Order[]
            {
                new Order { Quantity = 5 },
                new Order { Quantity = 10 },
                new Order { Quantity = 15 },
                new Order { Quantity = 20 },
                new Order { Quantity = 25 }
            };

            // Take a subset (indices 1 to 3 inclusive) using range syntax
            int[] quantitySlice = orders[1..4].Select(o => o.Quantity).ToArray();

            // Import the sliced quantities vertically starting at cell B2 (row 1, column 1)
            sheet.Cells.ImportArray(quantitySlice, 1, 1, true);

            // Save the workbook
            string outputPath = "SubsetImportDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO representing an order
    private class Order
    {
        public int Quantity { get; set; }
    }
}

public class Program
{
    public static void Main()
    {
        SubsetImportDemo.Run();
    }
}
