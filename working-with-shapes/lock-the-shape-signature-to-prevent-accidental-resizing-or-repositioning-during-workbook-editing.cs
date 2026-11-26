// Title: Lock a rectangle shape used as a signature placeholder to prevent resizing, moving, and selection in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock shape from resizing in Excel | prevent moving rectangle shape in Excel using Aspose.Cells | set shape locked properties with ShapeLockType in Aspose.Cells | protect worksheet objects to enforce shape locks Aspose.Cells | create signature placeholder shape and lock it in Aspose.Cells .NET
// Tags: Aspose.Cells shape lock resizing | Aspose.Cells ShapeLockType usage | C# protect worksheet objects Aspose.Cells | Aspose.Cells rectangle shape signature placeholder | Aspose.Cells lock shape selection | Aspose.Cells workbook shape protection

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeSignatureLockDemo
{
    // The example creates a new workbook, adds a rectangle shape as a signature placeholder, styles it, applies ShapeLockType flags to block resizing, moving, and selection, protects the worksheet for objects, and saves the workbook as an .xlsx file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SignatureSheet";

                // Add a rectangle shape that will act as a signature placeholder
                // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
                // Placed at row 2, column 2 with a size of 200x80 pixels
                Shape signatureShape = sheet.Shapes.AddRectangle(2, 2, 0, 0, 200, 80);
                signatureShape.Name = "SignatureShape";
                signatureShape.AlternativeText = "Signature placeholder";

                // Optional visual styling so the user can see the placeholder
                signatureShape.Line.Weight = 1.5;
                signatureShape.Line.DashStyle = MsoLineDashStyle.Solid;
                // Note: LineFormat does not expose a Color property in some versions; this line is omitted for compatibility
                signatureShape.Fill.FillType = FillType.Solid;
                signatureShape.Fill.SolidFill.Color = Color.LightYellow;

                // Lock the shape to prevent resizing, moving, and selection changes
                signatureShape.SetLockedProperty(ShapeLockType.AdjustHandles, true); // Prevent resizing via handles
                signatureShape.SetLockedProperty(ShapeLockType.Selection, true);    // Prevent selection (optional)
                signatureShape.SetLockedProperty(ShapeLockType.Move, true);        // Prevent moving

                // Protect the worksheet for objects so that locked shapes cannot be edited
                sheet.Protect(ProtectionType.Objects);

                // Save the workbook
                string outputPath = "SignatureLockedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The signature shape is locked against resizing and repositioning.");
                Console.WriteLine("Open the file in Excel, try to move or resize the shape – it will be prevented.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
