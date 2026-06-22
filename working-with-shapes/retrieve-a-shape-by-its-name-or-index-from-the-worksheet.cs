using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some shapes to the worksheet
        // Rectangle will be named "Rectangle 1" by default
        worksheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);
        // Oval will be named "Oval 1" by default
        worksheet.Shapes.AddOval(2, 0, 150, 100, 100, 100);

        // Retrieve a shape by its zero‑based index
        Shape shapeByIndex = worksheet.Shapes[0];
        Console.WriteLine($"Shape at index 0: Type = {shapeByIndex.Type}, Name = {shapeByIndex.Name}");

        // Retrieve a shape by its name
        Shape shapeByName = worksheet.Shapes["Rectangle 1"];
        if (shapeByName != null)
        {
            Console.WriteLine($"Shape named 'Rectangle 1' found: Type = {shapeByName.Type}");
        }
        else
        {
            Console.WriteLine("Shape named 'Rectangle 1' not found.");
        }

        // Save the workbook
        workbook.Save("RetrieveShapeDemo.xlsx");
    }
}