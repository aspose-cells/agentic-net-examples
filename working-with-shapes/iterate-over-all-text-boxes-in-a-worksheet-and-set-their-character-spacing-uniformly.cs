// Title: Increase the font size of every TextBox in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to increase font size of all text boxes in an Excel file using Aspose.Cells C# | Aspose.Cells C# iterate through worksheet TextBox objects and modify Font.Size | bulk change TextBox font size in a workbook with Aspose.Cells .NET API | C# code to loop over Excel text boxes and adjust their font size using Aspose.Cells | Aspose.Cells example for updating TextBox font properties across a worksheet
// Tags: Aspose.Cells TextBox Font.Size modification | C# iterate worksheet TextBox objects | bulk update Excel TextBox fonts Aspose.Cells | Aspose.Cells drawing objects font adjustment | increase TextBox font size Excel .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, iterates over each TextBox on the first worksheet, increments its Font.Size by a specified value, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or any specific worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the increment for font size (Aspose.Cells.Drawing.Font.Size is an int)
                int fontSizeIncrement = 2; // Adjust as needed

                // Iterate over all text boxes in the worksheet
                foreach (TextBox textBox in worksheet.TextBoxes)
                {
                    // Increase the font size to illustrate a change
                    textBox.Font.Size += fontSizeIncrement;
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
