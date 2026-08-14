// Title: C# LINQ Example to Disable RotateTextWithShape for Shapes in Aspose.Cells .NET
// Description: A concise Aspose.Cells for .NET demo that creates a workbook, adds a textbox and a rectangle, then uses LINQ to select every shape whose TextBody.TextAlignment.RotateTextWithShape flag is true and switches it off before saving the file. Shows how to batch‑disable text rotation that follows shape rotation.
// Keywords: Aspose.Cells | C# | .NET | RotateTextWithShape | disable RotateTextWithShape | LINQ shape filter | shape text rotation | Aspose.Cells shape properties | batch update shapes | worksheet shapes | text box rotation | Aspose.Cells API example
// Common Searches: how to turn off RotateTextWithShape in Aspose.Cells | LINQ query to find shapes with RotateTextWithShape enabled | disable text rotation with shape Aspose.Cells C# | batch modify shape properties Aspose.Cells workbook | Aspose.Cells example for rotating text boxes
// Developer Intent: Locate all shapes where RotateTextWithShape is true and set the property to false using Aspose.Cells.
// Use Cases: Iterate through a worksheet’s text boxes and ensure text stays horizontal after rotating the shapes. | Process a template workbook to remove linked text rotation before publishing. | Apply a single LINQ pass to update shape properties across large spreadsheets for performance. | Automate cleanup of imported drawings where RotateTextWithShape was unintentionally enabled.
// AI Prompts: Write C# code with Aspose.Cells that finds shapes with RotateTextWithShape enabled and disables it using LINQ. | Show how to batch‑update the TextBody.TextAlignment.RotateTextWithShape flag for multiple shapes in a workbook. | Explain best practices for safely changing RotateTextWithShape on a collection of shapes before saving.

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A concise Aspose.Cells for .NET demo that creates a workbook, adds a textbox and a rectangle, then uses LINQ to select every shape whose TextBody.TextAlignment.RotateTextWithShape flag is true and switches it off before saving the file. Shows how to batch‑disable text rotation that follows shape rotation.
public class DisableRotateTextWithShapeDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes for demonstration
        Shape shape1 = worksheet.Shapes.AddTextBox(1, 0, 1, 0, 100, 200);
        shape1.Text = "Rotated Text 1";
        shape1.TextBody.TextAlignment.RotateTextWithShape = true; // enable rotation with shape
        shape1.RotationAngle = 45; // rotate the shape itself

        Shape shape2 = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 250);
        shape2.Text = "Normal Text";
        shape2.TextBody.TextAlignment.RotateTextWithShape = false; // no rotation with shape

        // Use LINQ to select all shapes where RotateTextWithShape is true
        var shapesToUpdate = worksheet.Shapes
            .Cast<Shape>()
            .Where(s => s.TextBody != null && s.TextBody.TextAlignment.RotateTextWithShape);

        // Disable rotation with shape for each selected shape
        foreach (var shp in shapesToUpdate)
        {
            shp.TextBody.TextAlignment.RotateTextWithShape = false;
        }

        // Save the workbook
        string outputPath = "DisableRotateTextWithShapeDemo.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to {outputPath}");
    }
}
