// Title: Reusable C# method to add a custom property to an Excel worksheet using Aspose.Cells
// Common Searches: how to add a custom worksheet property with Aspose.Cells in C# | Aspose.Cells reusable helper for adding custom properties to a worksheet | C# example of adding a custom property to an Excel workbook using Aspose.Cells | add custom property to first worksheet Aspose.Cells .NET | ExcelHelper class for worksheet custom properties Aspose.Cells
// Tags: add custom worksheet property Aspose.Cells | Aspose.Cells reusable property helper C# | worksheet custom properties .NET | Excel workbook add custom property C# | Aspose.Cells custom property utility

using System;
using System.IO;
using Aspose.Cells;

// The example defines a static ExcelHelper class with an AddCustomProperty method that inserts a custom property into a worksheet's CustomProperties collection. It shows how to create a new Workbook, retrieve the first worksheet, call the helper to add a property, and save the workbook to an output file.
public static class ExcelHelper
{
    // Adds a custom property to the specified worksheet.
    public static void AddCustomProperty(Worksheet worksheet, string propertyName, string propertyValue)
    {
        // Add the property to the worksheet's custom properties collection.
        worksheet.CustomProperties.Add(propertyName, propertyValue);
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed).
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a custom property to the worksheet.
            ExcelHelper.AddCustomProperty(sheet, "MyProperty", "SomeValue");

            // Define output path.
            string outputPath = "output.xlsx";

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
