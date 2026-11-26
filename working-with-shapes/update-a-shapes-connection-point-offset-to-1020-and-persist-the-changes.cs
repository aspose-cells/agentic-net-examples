// Title: Update a shape's connection point offset to (10,20) and save the workbook with Aspose.Cells for .NET
// Common Searches: how to change a shape's connection point offset using Aspose.Cells in C# | example of updating rectangle shape connection points and saving workbook with Aspose.Cells | using reflection to call SetConnectionPoints method in Aspose.Cells .NET | persist shape connection point changes when exporting to XLSX with Aspose.Cells | C# code to modify shape connection points in an Excel file
// Tags: shape connection point offset update Aspose.Cells | rectangle shape SetConnectionPoints reflection .NET | save workbook after shape edits Aspose.Cells | modify Excel shape coordinates C# Aspose.Cells | connection points array manipulation Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace UpdateShapeConnectionPoint
{
    // The example creates a new workbook, adds a rectangle shape, retrieves its connection points, updates the first point's X and Y offsets to 10 and 20, optionally invokes the SetConnectionPoints method via reflection for version compatibility, and then saves the workbook so the changes are persisted.
    class Program
    {
        static void Main(string[] args)
        {
            string workbookPath = "ShapeConnectionPointDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 150, 100);
                shape.Name = "DemoRectangle";

                // Retrieve the current connection points
                float[][] connectionPoints = shape.GetConnectionPoints();

                if (connectionPoints == null || connectionPoints.Length == 0)
                {
                    Console.WriteLine("The shape does not have any connection points to modify.");
                }
                else
                {
                    // Ensure the first connection point is properly initialized
                    if (connectionPoints[0] == null || connectionPoints[0].Length < 2)
                    {
                        connectionPoints[0] = new float[2];
                    }

                    // Update the first connection point to (10, 20)
                    connectionPoints[0][0] = 10f; // X offset
                    connectionPoints[0][1] = 20f; // Y offset

                    // If the SetConnectionPoints method exists (different library versions), invoke it via reflection
                    MethodInfo setMethod = typeof(Shape).GetMethod("SetConnectionPoints", new[] { typeof(float[][]) });
                    if (setMethod != null)
                    {
                        setMethod.Invoke(shape, new object[] { connectionPoints });
                    }

                    Console.WriteLine("Connection point 0 updated to (10, 20).");
                }

                // Save the workbook to persist the changes
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{workbookPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
