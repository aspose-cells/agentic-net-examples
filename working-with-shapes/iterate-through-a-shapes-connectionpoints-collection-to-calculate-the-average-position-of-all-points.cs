// Title: Calculate the average position of a shape's connection points using Aspose.Cells for .NET (C#)
// Common Searches: how to get average X and Y coordinates of a shape's connection points with Aspose.Cells C# | iterate over Shape.ConnectionPoints collection in Aspose.Cells .NET | retrieve connection point positions from a named shape in an Excel workbook using Aspose.Cells | handle RuntimeBinderException when accessing ConnectionPoints property in Aspose.Cells | compute centroid of shape connection points in Excel with Aspose.Cells API
// Tags: Aspose.Cells shape connection points enumeration | average coordinates calculation C# | dynamic ConnectionPoints access Aspose.Cells | runtime binder exception handling Aspose.Cells | Excel shape geometry processing Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, finds a shape named "MyShape", and uses dynamic typing to access its ConnectionPoints collection. It iterates through each connection point, sums the X and Y values from the PointF Position property, computes the average X and Y coordinates, and prints the result. The code also demonstrates handling RuntimeBinderException for versions that do not support ConnectionPoints and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Find the shape named "MyShape"
            Shape targetShape = null;
            foreach (Shape shp in worksheet.Shapes)
            {
                if (shp.Name == "MyShape")
                {
                    targetShape = shp;
                    break;
                }
            }

            if (targetShape == null)
            {
                Console.WriteLine("Shape 'MyShape' not found.");
            }
            else
            {
                // Attempt to work with connection points (may not be supported in older versions)
                try
                {
                    dynamic dynShape = targetShape;
                    dynamic connPoints = dynShape.ConnectionPoints;

                    if (connPoints != null && connPoints.Count > 0)
                    {
                        float sumX = 0f;
                        float sumY = 0f;
                        int pointCount = connPoints.Count;

                        foreach (dynamic cp in connPoints)
                        {
                            // Each connection point is expected to have a Position property of type PointF
                            PointF pos = cp.Position;
                            sumX += pos.X;
                            sumY += pos.Y;
                        }

                        float avgX = sumX / pointCount;
                        float avgY = sumY / pointCount;
                        Console.WriteLine($"Average Position of Connection Points: X = {avgX}, Y = {avgY}");
                    }
                    else
                    {
                        Console.WriteLine("Shape has no connection points.");
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    Console.WriteLine("Connection points are not supported in the current Aspose.Cells version.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while processing connection points: {ex.Message}");
                }
            }

            // Save the workbook (optional, if any changes were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
