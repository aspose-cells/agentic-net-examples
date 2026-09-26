// Title: Load an Excel workbook from a byte array, replace a rectangle shape, and return the updated file as a byte array using Aspose.Cells for .NET
// AI Prompts: Write a C# method that takes a byte[] containing an XLSX workbook, deletes any shape named "MyRectangle", inserts a new rectangle with custom text and formatting, and returns the updated workbook as a byte[]. | Show how to open an Excel workbook from a MemoryStream using Aspose.Cells, shift the position of the first existing shape, and write the modified file to another MemoryStream. | Design a reusable routine that accepts parameters for rectangle size, fill color, and label text, updates the workbook's shapes in memory, and outputs the result as a byte array without touching the file system.
// Common Searches: Aspose.Cells C# load workbook from byte array and edit shapes in memory | How to delete a specific shape in an Excel file using Aspose.Cells without writing to disk | Saving modified Excel workbook to a byte[] with Aspose.Cells .NET | Replace rectangle shape and adjust shape position using Aspose.Cells C# example
// Tags: load workbook from memory stream Aspose.Cells | remove shape programmatically Aspose.Cells | add rectangle shape Aspose.Cells | export workbook to byte array Aspose.Cells | configure shape formatting Aspose.Cells C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates loading an XLSX workbook from a byte array via a MemoryStream, removing any existing shape named "MyRectangle", adding a new formatted rectangle shape, optionally repositioning the first shape, and saving the modified workbook back to a byte array—all without accessing the file system.
public class WorkbookShapeProcessor
{
    /// <param name="inputWorkbookBytes">The original workbook data.</param>
    /// <returns>The modified workbook data.</returns>
    public byte[] ProcessWorkbook(byte[] inputWorkbookBytes)
    {
        try
        {
            // Load the workbook from the input memory stream
            using (MemoryStream inputStream = new MemoryStream(inputWorkbookBytes))
            {
                Workbook workbook = new Workbook(inputStream);
                Worksheet sheet = workbook.Worksheets[0];

                // Remove existing shapes named "MyRectangle" to avoid duplicates
                for (int i = sheet.Shapes.Count - 1; i >= 0; i--)
                {
                    Shape existingShape = sheet.Shapes[i];
                    if (existingShape.Name == "MyRectangle")
                    {
                        sheet.Shapes.RemoveAt(i);
                    }
                }

                // Add a new rectangle shape
                Shape rectangle = sheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 100);
                rectangle.Name = "MyRectangle";
                rectangle.Text = "Hello Aspose.Cells!";
                rectangle.Font.Color = Color.White;
                rectangle.FillFormat.ForeColor = Color.Blue;
                rectangle.LineFormat.ForeColor = Color.Black;
                rectangle.LineFormat.Weight = 1.5;

                // Move the first existing shape (if any)
                if (sheet.Shapes.Count > 0)
                {
                    Shape firstShape = sheet.Shapes[0];
                    firstShape.Left += 10;
                    firstShape.Top += 5;
                }

                // Save the modified workbook to an output memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    return outputStream.ToArray();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string inputPath = "input.xlsx";
            byte[] inputBytes;

            if (File.Exists(inputPath))
            {
                inputBytes = File.ReadAllBytes(inputPath);
            }
            else
            {
                // Create a new empty workbook if the input file does not exist
                Workbook wb = new Workbook();
                using (MemoryStream ms = new MemoryStream())
                {
                    wb.Save(ms, SaveFormat.Xlsx);
                    inputBytes = ms.ToArray();
                }
            }

            WorkbookShapeProcessor processor = new WorkbookShapeProcessor();
            byte[] resultBytes = processor.ProcessWorkbook(inputBytes);

            string outputPath = "output.xlsx";
            File.WriteAllBytes(outputPath, resultBytes);
            Console.WriteLine($"Modified workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
