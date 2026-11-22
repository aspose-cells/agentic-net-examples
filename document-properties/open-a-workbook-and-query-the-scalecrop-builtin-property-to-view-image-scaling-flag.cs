// Title: Read the ScaleCrop image scaling flag of a picture in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get picture ScaleCrop flag from Excel file | read image scaling setting (ScaleCrop) of embedded picture with Aspose.Cells | using reflection to access ScaleCrop property of a picture in Aspose.Cells | check if picture is set to scale and crop in Excel via Aspose.Cells .NET | retrieve built‑in ScaleCrop attribute of worksheet picture using Aspose.Cells API
// Tags: Aspose.Cells picture ScaleCrop retrieval | C# read Excel picture scaling flag | reflection access built-in image property Aspose.Cells | worksheet picture collection handling Aspose.Cells | ScaleCrop property usage Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook, accesses the first worksheet's picture collection, and uses reflection to read the boolean ScaleCrop flag of the first picture, handling cases where the property is unavailable and outputting the result.
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

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust the index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Access the collection of pictures on the worksheet
            PictureCollection pictures = sheet.Pictures;

            if (pictures.Count > 0)
            {
                // Retrieve the first picture
                Picture picture = pictures[0];

                // Attempt to read the ScaleCrop flag via reflection (property may not exist in older versions)
                bool scaleCropFlag = false;
                try
                {
                    PropertyInfo propInfo = picture.GetType().GetProperty("ScaleCrop", BindingFlags.Public | BindingFlags.Instance);
                    if (propInfo != null && propInfo.PropertyType == typeof(bool))
                    {
                        scaleCropFlag = (bool)propInfo.GetValue(picture);
                        Console.WriteLine("ScaleCrop flag: " + scaleCropFlag);
                    }
                    else
                    {
                        Console.WriteLine("ScaleCrop property is not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error retrieving ScaleCrop flag: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("No pictures found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
