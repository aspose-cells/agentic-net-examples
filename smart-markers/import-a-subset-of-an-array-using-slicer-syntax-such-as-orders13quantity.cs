// Title: Import a subset of an Order array into an Excel worksheet with Aspose.Cells using a SUM formula as a slice‑syntax workaround (C#)
// Common Searches: how to import only part of a C# array into Aspose.Cells worksheet | Aspose.Cells example for summing selected items of a POCO collection | C# workaround for missing slice syntax when using Aspose.Cells smart markers | using Excel SUM formula in Aspose.Cells to total specific rows | extracting array slice 1..3 to Excel cells with Aspose.Cells C#
// Tags: Aspose.Cells import array subset C# | Excel SUM formula with Aspose.Cells workbook | C# POCO collection to Excel range Aspose.Cells | slice syntax alternative Aspose.Cells | calculate selected quantities Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsObjectModelExample
{
    // Simple POCO representing an order
    // The sample creates an Order[] array, writes each Quantity to cells A1‑A4 of a new workbook, applies the Excel formula =SUM(A2:A4) to cell A5 to sum the 2nd‑4th orders (90), calculates the formula, outputs the result, and saves the file as ObjectModelSliceExample.xlsx, demonstrating a workaround for the missing slice‑syntax feature in Aspose.Cells.
    public class Order
    {
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Prepare sample data: an array of Order objects
                Order[] orders = new Order[]
                {
                    new Order { Quantity = 10 },
                    new Order { Quantity = 20 },
                    new Order { Quantity = 30 },
                    new Order { Quantity = 40 }
                };

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate the worksheet with the order quantities (A1:A4)
                for (int i = 0; i < orders.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(orders[i].Quantity);
                }

                // Use a regular Excel formula to sum the quantities of the 2nd, 3rd, and 4th orders
                // (cells A2:A4). This replaces the object‑model slice syntax which is unavailable
                // in the current Aspose.Cells version.
                sheet.Cells["A5"].Formula = "=SUM(A2:A4)";

                // Calculate the formula
                workbook.CalculateFormula();

                // Retrieve and display the result (should be 20 + 30 + 40 = 90)
                Console.WriteLine("Sum of selected quantities: " + sheet.Cells["A5"].Value);

                // Save the workbook to a file
                string outputPath = "ObjectModelSliceExample.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
