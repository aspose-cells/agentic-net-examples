// Title: Read a shape's shadow color and log its RGB values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get the shadow color of a shape in an Excel worksheet | How to read RGB components of a shape's shadow effect with Aspose.Cells for .NET | Programmatically retrieve a shape's shadow color from a workbook using Aspose.Cells | C# example to log the ARGB value of a shape's shadow in Excel using Aspose.Cells | Extract shadow effect color from an existing Excel file with Aspose.Cells API
// Tags: shape shadow effect extraction Aspose.Cells C# | retrieve rgb from CellsColor shadow Aspose.Cells | log shape shadow ARGB Aspose.Cells .NET | access ShadowEffect property Excel shape C# | read shape shadow properties from workbook Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShadowColorReader
{
    // The sample loads (or creates) an Excel workbook, adds a rectangle shape with a SteelBlue shadow, accesses the shape's ShadowEffect, obtains its CellsColor, converts it to System.Drawing.Color, and writes both the ARGB hex value and the individual R, G, B components to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the workbook file to work with.
            // If a path is supplied as a command‑line argument, try to load it.
            // Otherwise, create a new workbook with a shape that has a known shadow color.
            string workbookPath = args.Length > 0 ? args[0] : "ShadowDemo.xlsx";

            Workbook workbook = null;
            bool createdTempWorkbook = false;

            try
            {
                if (System.IO.File.Exists(workbookPath))
                {
                    Console.WriteLine($"Loading existing workbook: {workbookPath}");
                    workbook = new Workbook(workbookPath);
                }
                else
                {
                    Console.WriteLine($"Workbook not found. Creating a new workbook and saving to: {workbookPath}");
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];

                    // Add a rectangle shape.
                    Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

                    // Ensure the shape has a shadow effect and set a known color.
                    ShadowEffect shadow = shape.ShadowEffect;
                    CellsColor cellsColor = workbook.CreateCellsColor();
                    cellsColor.Color = Color.FromArgb(255, 70, 130, 180); // SteelBlue
                    shadow.Color = cellsColor;

                    // Save the workbook so that the shadow color persists.
                    workbook.Save(workbookPath);
                    createdTempWorkbook = true;
                    Console.WriteLine("Sample workbook created with a shape that has a shadow color.");
                }

                // Retrieve the first shape from the first worksheet.
                Worksheet firstSheet = workbook.Worksheets[0];
                if (firstSheet.Shapes.Count == 0)
                {
                    Console.WriteLine("No shapes found in the worksheet. Exiting.");
                    return;
                }

                Shape targetShape = firstSheet.Shapes[0];
                Console.WriteLine($"Processing shape: Type={targetShape.AutoShapeType}, Name={targetShape.Name}");

                // Access the shadow effect.
                ShadowEffect shadowEffect = targetShape.ShadowEffect;
                if (shadowEffect == null)
                {
                    Console.WriteLine("The shape does not have a ShadowEffect object.");
                    return;
                }

                // Retrieve the shadow color.
                CellsColor shadowCellsColor = shadowEffect.Color;
                if (shadowCellsColor == null)
                {
                    Console.WriteLine("Shadow color is not set (null).");
                    return;
                }

                // Convert to System.Drawing.Color to obtain RGB components.
                Color sysColor = shadowCellsColor.Color;
                Console.WriteLine($"Shadow Color (ARGB): 0x{shadowCellsColor.Argb:X8}");
                Console.WriteLine($"Shadow Color (RGB): R={sysColor.R}, G={sysColor.G}, B={sysColor.B}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // If we created a temporary workbook for demonstration and the user did not request it,
                // we leave the file on disk so they can inspect it later.
                if (createdTempWorkbook)
                {
                    Console.WriteLine($"Demo workbook saved at: {System.IO.Path.GetFullPath(workbookPath)}");
                }
            }
        }
    }
}
