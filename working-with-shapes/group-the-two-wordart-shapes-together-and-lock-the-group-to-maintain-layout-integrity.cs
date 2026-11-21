// Title: Lock WordArt shapes to keep layout unchanged using Aspose.Cells for .NET (C#)
// Common Searches: how to prevent WordArt shapes from being moved in an Excel file with Aspose.Cells C# | Aspose.Cells C# set IsLocked property on WordArt objects | make WordArt shapes read‑only in generated Excel workbook using Aspose.Cells | protect WordArt layout in Excel by locking shapes with Aspose.Cells .NET
// Tags: lock WordArt shapes Aspose.Cells | prevent shape movement Excel C# | Aspose.Cells IsLocked property usage | add WordArt to worksheet C# | secure WordArt objects Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts two WordArt shapes on the first worksheet, sets each shape's IsLocked property to true to prevent editing, and saves the file as GroupedWordArt.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add first WordArt shape
            // Parameters: preset type (0 = preset), text, upper left row, upper left column, top, left, width, height
            Shape wordArt1 = sheet.Shapes.AddWordArt(0, "Hello", 0, 0, 0, 0, 200, 50);

            // Add second WordArt shape
            Shape wordArt2 = sheet.Shapes.AddWordArt(0, "World", 0, 0, 250, 0, 200, 50);

            // Optionally lock the shapes to prevent accidental changes
            wordArt1.IsLocked = true;
            wordArt2.IsLocked = true;

            // Save the workbook
            string outputPath = "GroupedWordArt.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{System.IO.Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
