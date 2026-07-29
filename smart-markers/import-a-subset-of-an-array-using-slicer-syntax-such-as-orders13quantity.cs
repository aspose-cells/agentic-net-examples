// Title: Import a Subset of an Object Array into Excel with Aspose.Cells using C# Range Operator
// Description: Demonstrates how to slice an Order[] array with the C# range syntax (orders[1..4]), extract the Quantity values, and write them vertically to cell B2 using Aspose.Cells Workbook.ImportArray, then save the workbook as ImportSubsetDemo.xlsx.
// Keywords: Aspose.Cells ImportArray | C# range operator | array slicing Excel | .NET Excel export | smart markers partial import | ImportArray int[] | Excel data subset C#
// Common Searches: Aspose.Cells import part of array | C# range syntax with ImportArray | how to write selected rows to Excel using Aspose | slice object array and export to Excel .NET | smart markers import subset example
// Developer Intent: Select specific elements from an object collection and export only those values to an Excel worksheet with Aspose.Cells.
// Use Cases: Generate a sales sheet that includes only orders 2‑4 by importing their quantities. | Create a report that shows the most recent entries from a larger dataset without processing the entire collection. | Build a dynamic workbook that writes a filtered range of data to a predefined cell range.
// AI Prompts: Show C# code that uses the range operator to slice an array and imports the slice into Excel with Aspose.Cells ImportArray. | Explain how to extract a property from a sliced object array and write it vertically starting at cell B2. | Provide guidance on adjusting range indices to include the desired elements when using ImportArray.

using Aspose.Cells;
using System;

namespace Demo
{
    // Simple data model representing an order with a quantity.
    // Demonstrates how to slice an Order[] array with the C# range syntax (orders[1..4]), extract the Quantity values, and write them vertically to cell B2 using Aspose.Cells Workbook.ImportArray, then save the workbook as ImportSubsetDemo.xlsx.
    public class Order
    {
        public int Quantity { get; set; }
        public Order(int qty) => Quantity = qty;
    }

    // Demonstrates importing a subset of data into an Excel worksheet.
    public class ImportSubsetDemo
    {
        public static void Run()
        {
            // Sample data: array of Order objects.
            Order[] orders = new Order[]
            {
                new Order(10),
                new Order(20),
                new Order(30),
                new Order(40),
                new Order(50)
            };

            // Use C# range syntax to select a subset (indexes 1 to 3 inclusive).
            // Range end is exclusive, so use 1..4 to include index 3.
            Order[] subset = orders[1..4];

            // Extract the Quantity property values into a simple int array.
            int[] quantities = new int[subset.Length];
            for (int i = 0; i < subset.Length; i++)
            {
                quantities[i] = subset[i].Quantity;
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Import the quantities vertically starting at cell B2 (row 1, column 1 in zero‑based indices).
            sheet.Cells.ImportArray(quantities, 1, 1, true);

            // Save the workbook.
            workbook.Save("ImportSubsetDemo.xlsx");
        }
    }

    // Entry point for the console application.
    public class Program
    {
        public static void Main()
        {
            try
            {
                ImportSubsetDemo.Run();
                Console.WriteLine("Workbook saved successfully as ImportSubsetDemo.xlsx.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
