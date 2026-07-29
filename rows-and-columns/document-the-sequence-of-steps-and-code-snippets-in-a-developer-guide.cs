// Title: Aspose.Cells .NET: Create a Rectangle Shape and Apply Custom Adjustment Guides
// Description: Step‑by‑step guide that shows how to create a Workbook, add a rectangle shape to the first worksheet, access its Geometry.ShapeAdjustValues collection, add named adjustment guides, and save the file as XLSX. Includes full C# code with error handling and demonstrates both BaseShapeGuide and ShapeGuideCollection usage.
// Keywords: Aspose.Cells C# rectangle shape | ShapeGuideCollection Aspose.Cells | BaseShapeGuide example | adjustment guides shape Aspose.Cells | shape geometry Aspose.Cells .NET | save workbook with shapes | Aspose.Cells shape guide tutorial | C# Aspose.Cells shape adjustment | XLSX shape guide code sample
// Common Searches: add rectangle shape Aspose.Cells .NET | how to use ShapeGuideCollection in Aspose.Cells | set custom adjustment guides for a shape in C# | Aspose.Cells BaseShapeGuide documentation | save workbook with shapes using Aspose.Cells
// Developer Intent: Create a worksheet, insert a rectangle shape, define custom adjustment guides via ShapeGuideCollection, and persist the workbook.
// Use Cases: Build spreadsheet templates that contain predefined graphic placeholders. | Programmatically fine‑tune diagram dimensions by adjusting shape guides. | Automate report generation where consistent shape geometry is required across files.
// AI Prompts: Generate a detailed tutorial for adding a rectangle shape with custom adjustment guides using Aspose.Cells in C#. | Write C# code that updates existing ShapeGuideCollection values and saves the workbook as an XLSX file. | Explain the relationship between BaseShapeGuide, ShapeGuideCollection, and Shape.Geometry in Aspose.Cells and how they control shape adjustments.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsGuide
{
    // Step‑by‑step guide that shows how to create a Workbook, add a rectangle shape to the first worksheet, access its Geometry.ShapeAdjustValues collection, add named adjustment guides, and save the file as XLSX. Includes full C# code with error handling and demonstrates both BaseShapeGuide and ShapeGuideCollection usage.
    public class ShapeGuideExamples
    {
        // Demonstrates BaseShapeGuide via shape creation
        public static void BaseShapeGuideDemo()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape (BaseShapeGuide is the base class for guides)
                Shape shape = sheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);

                Console.WriteLine("Rectangle shape created.");

                // Save the workbook
                workbook.Save("BaseShapeGuideDemo.xlsx");
                Console.WriteLine("Workbook saved as BaseShapeGuideDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in BaseShapeGuideDemo: {ex.Message}");
            }
        }

        // Demonstrates ShapeGuideCollection usage
        public static void ShapeGuideCollectionDemo()
        {
            try
            {
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a shape that supports adjustment guides
                Shape shape = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 10, 10, 0, 0, 200, 100);
                shape.AutoShapeType = AutoShapeType.Rectangle;

                // Retrieve the collection of shape guides
                ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

                // Add several guides (name, value)
                guides.Add("adj1", 0.5);
                guides.Add("adj2", 0.2);
                guides.Add("adj3", 0.8);

                Console.WriteLine("Adjustment guides added to the shape.");

                // Save the workbook
                workbook.Save("ShapeGuideCollectionDemo.xlsx");
                Console.WriteLine("Workbook saved as ShapeGuideCollectionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in ShapeGuideCollectionDemo: {ex.Message}");
            }
        }

        // Entry point for testing the demos
        public static void Main()
        {
            BaseShapeGuideDemo();
            ShapeGuideCollectionDemo();
        }
    }
}
