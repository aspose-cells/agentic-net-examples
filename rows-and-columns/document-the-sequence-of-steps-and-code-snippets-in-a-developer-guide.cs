using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Vba;

public class ShapeGuideDeveloperGuide
{
    // Demonstrates creating a shape and accessing its BaseShapeGuide via Geometry.
    public static void BaseShapeGuideDemo()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape.
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 0, 100, 100, 100);

        // BaseShapeGuide has no public constructor; we work with ShapeGuide derived from it.
        Console.WriteLine($"Created shape of type: {shape.GetType().Name}");

        // Save the workbook.
        workbook.Save("BaseShapeGuideDemo.xlsx");
    }

    // Demonstrates adding and retrieving shape guides using ShapeGuideCollection.
    public static void ShapeGuideCollectionDemo()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a shape that supports adjustment guides.
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.RightArrowCallout, 2, 0, 2, 0, 200, 150);

        // Get the collection of guides.
        ShapeGuideCollection guides = shape.Geometry.ShapeAdjustValues;

        // Add guides with names and values.
        guides.Add("adj1", 25.5);
        guides.Add("adj2", 30);
        guides.Add("adj3", 25.5);
        guides.Add("adj4", 35);

        // Read and display guide values.
        for (int i = 0; i < guides.Count; i++)
        {
            Console.WriteLine($"Guide {i} value: {guides[i].Value}");
        }

        // Modify a guide value.
        guides[0].Value = 20.0;

        // Save the workbook.
        workbook.Save("ShapeGuideCollectionDemo.xlsx");
    }

    // Demonstrates creating a ShapeGuide via Add method and accessing its Value property.
    public static void ShapeGuideDemo()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add an auto shape.
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.RoundedRectangle, 0, 0, 100, 100, 100, 100);

        // Add a guide and obtain its index.
        int index = shape.Geometry.ShapeAdjustValues.Add("Guide1", 0.2);
        ShapeGuide guide = shape.Geometry.ShapeAdjustValues[index];

        Console.WriteLine($"Created ShapeGuide with initial value: {guide.Value}");

        // Update the guide value.
        guide.Value = 0.5;
        Console.WriteLine($"Updated ShapeGuide value: {guide.Value}");

        // Save the workbook.
        workbook.Save("ShapeGuideDemo.xlsx");
    }

    // Demonstrates using ShapePath to create a custom freeform shape.
    public static void ShapePathDemo()
    {
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a ShapePath instance and define a rectangle.
        ShapePath path = new ShapePath();
        path.MoveTo(100, 100);
        path.LineTo(200, 100);
        path.LineTo(200, 200);
        path.LineTo(100, 200);
        path.Close();

        // Add the freeform shape using the defined path.
        sheet.Shapes.AddFreeform(1, 0, 1, 0, 200, 200, new ShapePath[] { path });

        // Save the workbook.
        workbook.Save("ShapePathDemo.xlsx");
    }

    // Demonstrates adding a VBA module and setting its code.
    public static void VbaModuleDemo()
    {
        Workbook workbook = new Workbook();
        VbaProject vbaProject = workbook.VbaProject;

        // Add a class module named "MyClass".
        int idx = vbaProject.Modules.Add(VbaModuleType.Class, "MyClass");
        VbaModule module = vbaProject.Modules[idx];
        module.Codes = "Sub ShowMessage()\r\nMsgBox \"Hello from VBA\"\r\nEnd Sub";

        // Save the workbook with VBA project.
        workbook.Save("VbaModuleDemo.xlsm");
    }

    // Runs all demonstration methods sequentially.
    public static void RunAll()
    {
        BaseShapeGuideDemo();
        ShapeGuideCollectionDemo();
        ShapeGuideDemo();
        ShapePathDemo();
        VbaModuleDemo();
    }
}

public class Program
{
    public static void Main()
    {
        ShapeGuideDeveloperGuide.RunAll();
    }
}