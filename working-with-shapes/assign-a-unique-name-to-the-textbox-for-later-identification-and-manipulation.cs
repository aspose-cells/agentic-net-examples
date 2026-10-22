// Title: How to assign a unique name to a TextBox shape in an Aspose.Cells workbook using C#
// AI Prompts: Create a new Excel workbook, add a TextBox shape at row 3 column 2, assign a unique identifier via the Name property, and save the file using Aspose.Cells for .NET. | Generate C# code that adds a TextBox to a worksheet, assigns a unique name for later retrieval, and writes sample text inside the shape. | Write a method that locates a TextBox shape by its Name property in an existing workbook and updates its content using Aspose.Cells.
// Common Searches: Aspose.Cells C# set custom name for textbox shape in Excel worksheet | How to rename a TextBox shape added with Aspose.Cells .NET | Identify a specific textbox in an Excel file using Aspose.Cells Name property | C# example for adding a named textbox to a workbook with Aspose.Cells
// Tags: Aspose.Cells set TextBox.Name property | C# add textbox shape with custom identifier to Excel | Aspose.Cells retrieve shape using Name property | Excel workbook textbox naming Aspose.Cells | Aspose.Cells shape manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example creates a new workbook, inserts a TextBox shape on the first worksheet, assigns it the unique name "UniqueTextBox_001", sets sample text, and saves the workbook as Result.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox shape to the worksheet (row, column, offsetX, offsetY, width, height)
                // In recent Aspose.Cells versions AddTextBox returns the TextBox object directly
                TextBox textBox = worksheet.Shapes.AddTextBox(2, 1, 0, 0, 200, 100);

                // Assign a unique name to the textbox for later identification
                textBox.Name = "UniqueTextBox_001";

                // (Optional) Set some sample text inside the textbox
                textBox.Text = "Hello, Aspose!";

                // Define output file path
                string outputPath = "Result.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
