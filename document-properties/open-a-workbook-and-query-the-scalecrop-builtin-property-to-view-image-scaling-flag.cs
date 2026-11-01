// Title: Reading the ScaleCrop flag of picture shapes in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get ScaleCrop property of picture shape in Excel file | read image scaling flag from worksheet shapes using Aspose.Cells | use reflection to access ScaleCrop of pictures in an Aspose.Cells workbook | determine if a picture is set to scale and crop in a .xlsx with Aspose.Cells .NET | list ScaleCrop values for all pictures in an Excel sheet via Aspose.Cells API
// Tags: Aspose.Cells picture ScaleCrop retrieval | iterate worksheet shapes Aspose.Cells C# | reflection access Aspose.Cells picture property | read picture scaling flag .xlsx Aspose.Cells | image scaling attribute Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Reflection;

// The example loads an existing Excel workbook, accesses the first worksheet, iterates through its shapes, identifies picture objects, and uses reflection to read the boolean ScaleCrop property of each picture, outputting the flag value or indicating it is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Determine if the shape is a picture
                if (shape is Picture picture)
                {
                    bool? isScaleCrop = null;

                    try
                    {
                        // Attempt to read the ScaleCrop property via reflection
                        PropertyInfo propInfo = picture.GetType().GetProperty("ScaleCrop", BindingFlags.Public | BindingFlags.Instance);
                        if (propInfo != null && propInfo.PropertyType == typeof(bool))
                        {
                            isScaleCrop = (bool)propInfo.GetValue(picture);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log any reflection errors but continue processing
                        Console.WriteLine($"Warning: Unable to retrieve ScaleCrop for picture \"{picture.Name}\": {ex.Message}");
                    }

                    // Output the result (or N/A if property not available)
                    Console.WriteLine($"Picture \"{picture.Name}\" ScaleCrop flag: {(isScaleCrop.HasValue ? isScaleCrop.Value.ToString() : "N/A")}");
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
