// Title: How to write a one‑dimensional string array into the first row (A1 onward) of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through a string[] and uses Aspose.Cells Cells.PutValue to insert each element horizontally starting at cell A1. | Show the modification needed to fill the same string array vertically down column A with Aspose.Cells in C#. | Provide a C# snippet that saves the workbook to a user‑specified path after importing the string array with Aspose.Cells.
// Common Searches: Aspose.Cells C# write string array to first row starting at A1 | populate Excel cells horizontally from a one‑dimensional array using Aspose.Cells .NET | how to save an Aspose.Cells workbook after loading data from a string[] | C# loop to put values from string[] into worksheet cells with Aspose.Cells
// Tags: Aspose.Cells putvalue string array horizontal | import one-dimensional array into Excel worksheet .NET | save workbook after data import Aspose.Cells | C# loop fill cells from array Aspose.Cells | write array to first row A1 Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates an in‑memory Workbook, accesses the first Worksheet, defines a string[] of fruit names, and uses a for‑loop with Cells[0,i].PutValue to write each element into consecutive cells of the first row starting at A1. The workbook is then saved as Output.xlsx and the full file path is printed, with exception handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            // One‑dimensional string array to be loaded into the worksheet
            string[] data = new string[] { "Apple", "Banana", "Cherry", "Date" };

            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Import the array starting at row 0, column 0 (cell A1) horizontally
            // Using a loop for compatibility with all Aspose.Cells versions
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[0, i].PutValue(data[i]);
            }

            // Define output file path
            string outputPath = "Output.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
