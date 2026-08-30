// Title: Define a custom DisplayRange and loop through only the cells inside that visible region with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, set a specific display range such as B2:C3, and use the range enumerator to print each cell's address and value in C# with Aspose.Cells. | After establishing a custom display range, obtain the worksheet's MaxDisplayRange property and output its address to verify the overall visible area.
// Common Searches: C# Aspose.Cells how to enumerate cells only within a defined display range | example of using MaxDisplayRange property to get worksheet visible area in Aspose.Cells | Aspose.Cells create range B2:C3 and loop through cells in .NET | filter cells by visible region using DisplayRange in Aspose.Cells for C# | retrieve address of maximal display range after setting custom range Aspose.Cells
// Tags: custom display range Aspose.Cells C# | enumerate cells in specific range Aspose.Cells | MaxDisplayRange property usage Aspose.Cells | loop through visible cells worksheet Aspose.Cells | create range B2:C3 Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, defines a custom display range (B2:C3), enumerates and prints each cell's address and value within that range, retrieves the worksheet's MaxDisplayRange address, and saves the file as DisplayRangeDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample data
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue(100);
            cells["B2"].PutValue(200);
            cells["C2"].PutValue(300);
            cells["A3"].PutValue(400);
            cells["B3"].PutValue(500);
            cells["C3"].PutValue(600);

            // Define a specific display area (e.g., B2:C3)
            // This range will act as the "visible" region we want to work with
            AsposeRange displayRange = cells.CreateRange("B2", "C3");

            // Enumerate only the cells that belong to the defined display range
            IEnumerator enumerator = displayRange.GetEnumerator();
            while (enumerator.MoveNext())
            {
                Cell cell = (Cell)enumerator.Current;
                Console.WriteLine($"{cell.Name}: {cell.Value}");
            }

            // Optionally, retrieve the worksheet's maximal display range (includes data, merged cells, shapes)
            AsposeRange maxDisplayRange = worksheet.Cells.MaxDisplayRange;
            if (maxDisplayRange != null)
            {
                Console.WriteLine($"Max Display Range Address: {maxDisplayRange.Address}");
            }

            // Save the workbook
            string outputPath = "DisplayRangeDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
