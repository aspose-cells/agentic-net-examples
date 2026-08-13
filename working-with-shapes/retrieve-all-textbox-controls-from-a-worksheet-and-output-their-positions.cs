// Title: C# – Retrieve All TextBox Positions from an Excel Worksheet with Aspose.Cells
// Description: Loads an Excel workbook, accesses the first worksheet’s TextBoxCollection, iterates each TextBox, and prints its Top, Left, Height, and Width values to the console. The workbook can then be saved optionally.
// Keywords: Aspose.Cells | C# | Excel TextBox positions | shape coordinates | retrieve textbox size | list textbox locations | worksheet shapes | .NET Excel API
// Common Searches: Aspose.Cells get textbox coordinates | list all textbox positions in Excel using C# | how to read textbox size with Aspose.Cells | enumerate shape locations Aspose.Cells .NET | retrieve textbox top left values Aspose.Cells
// Developer Intent: The developer wants to enumerate every TextBox shape in a worksheet and obtain its exact location and dimensions.
// Use Cases: Generate a layout audit that records the position and size of each TextBox in a spreadsheet. | Programmatically shift TextBox shapes by calculated offsets. | Export textbox coordinates to an external system for visualization or further processing.
// AI Prompts: Write C# code that moves each TextBox in a worksheet 10 points down using Aspose.Cells. | Create a method that returns a list of objects containing Top, Left, Height, and Width for all TextBoxes in a given worksheet. | Explain how to filter TextBox shapes by size before retrieving their positions with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, accesses the first worksheet’s TextBoxCollection, iterates each TextBox, and prints its Top, Left, Height, and Width values to the console. The workbook can then be saved optionally.
    public class RetrieveTextBoxPositions
    {
        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the collection of TextBox objects
                TextBoxCollection textBoxes = worksheet.TextBoxes;

                // Iterate through all TextBoxes and output their positions
                for (int i = 0; i < textBoxes.Count; i++)
                {
                    TextBox tb = textBoxes[i];
                    Console.WriteLine($"TextBox {i}: Top={tb.Top}, Left={tb.Left}, Height={tb.Height}, Width={tb.Width}");
                }

                // Save the workbook (optional, can be the same file or a new one)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RetrieveTextBoxPositions.Run();
        }
    }
}
