// Title: C# – Add a Rectangle Shape, Assign a Unique Name, and Retrieve It with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert a rectangle shape into the first worksheet, set the shape's Name property to a custom identifier, retrieve the shape using the name indexer, and save the file as an XLSX document.
// Keywords: Aspose.Cells C# shape example | add rectangle shape Aspose.Cells | shape Name property .NET | retrieve shape by name Aspose.Cells | worksheet.Shapes indexer | Aspose.Cells sample code | C# Excel shape naming
// Common Searches: how to add a shape with a custom name using Aspose.Cells for .NET | retrieve worksheet shape by its Name property C# | Aspose.Cells set and get shape name | example of shape naming in Aspose.Cells | C# Aspose.Cells shape indexer by name
// Developer Intent: Create a shape, give it a unique Name, and later access the same shape via that name.
// Use Cases: Insert a labeled rectangle, assign a Name, and later modify its size or formatting by retrieving it with worksheet.Shapes["MyUniqueShape"] | Add multiple diagram elements with distinct names, then iterate over the names to update each shape dynamically | Map shape names to business data in a dictionary for quick visual updates in the workbook
// AI Prompts: Generate C# code that adds several shapes to a worksheet, assigns each a unique Name, and changes the fill color of a specific shape retrieved by its name using Aspose.Cells. | Explain how to safely handle exceptions when a shape name does not exist in a worksheet with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for Shape class

namespace AsposeCellsExample
{
    // Demonstrates how to create a new Workbook, insert a rectangle shape into the first worksheet, set the shape's Name property to a custom identifier, retrieve the shape using the name indexer, and save the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape: row, column, top offset, left offset, height, width
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 200);
                shape.Name = "MyUniqueShape";

                // Retrieve the shape by its unique name
                Shape retrievedShape = worksheet.Shapes["MyUniqueShape"];
                Console.WriteLine("Retrieved shape name: " + retrievedShape.Name);

                // Save the workbook
                string outputPath = "ShapeWithUniqueName.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
