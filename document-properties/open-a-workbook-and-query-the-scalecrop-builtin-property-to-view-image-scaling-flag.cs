// Title: Read the ScaleCrop flag of a picture in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Load an Excel file with Aspose.Cells, access the first worksheet picture, and output its ScaleCrop boolean value together with WidthScale and HeightScale. | Iterate over all pictures in a workbook and print each picture’s ScaleCrop status along with the corresponding WidthScale and HeightScale factors.
// Common Searches: Aspose.Cells get ScaleCrop property of worksheet picture C# | how to determine if an Excel image is scaled using Aspose.Cells .NET | C# read WidthScale and HeightScale of pictures in an Excel file with Aspose.Cells | retrieve picture scaling flag from Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells picture ScaleCrop property | read image scaling flag Aspose.Cells | C# query Excel picture dimensions | retrieve picture WidthScale HeightScale Aspose.Cells | check picture scaling status in worksheet .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The sample loads an Excel workbook, accesses the first worksheet, checks for pictures, and for each picture reads the ScaleCrop flag as well as the WidthScale and HeightScale properties to determine whether the image has been scaled, then prints the scaling status and scale factors.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Verify that at least one worksheet exists
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for pictures in the worksheet
            if (worksheet.Pictures.Count > 0)
            {
                // Get the first picture
                Picture picture = worksheet.Pictures[0];

                // Determine if the picture has been scaled (WidthScale or HeightScale differs from 1.0)
                bool isScaled = picture.WidthScale != 1.0 || picture.HeightScale != 1.0;

                Console.WriteLine($"Picture 0 scaled: {isScaled}");
                Console.WriteLine($"WidthScale: {picture.WidthScale}, HeightScale: {picture.HeightScale}");
            }
            else
            {
                Console.WriteLine("No pictures found in the worksheet.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
