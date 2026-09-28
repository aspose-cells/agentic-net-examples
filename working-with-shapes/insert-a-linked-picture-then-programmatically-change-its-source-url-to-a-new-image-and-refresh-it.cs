// Title: Insert a linked picture from a URL into an Excel worksheet and later change its source using Aspose.Cells for .NET (C#)
// AI Prompts: Create a linked picture in a worksheet cell that points to an external image URL with Aspose.Cells for .NET. | Update the URL of an existing linked picture in a workbook and refresh the image programmatically.
// Common Searches: how to add a web-linked image to a specific cell in Excel using Aspose.Cells C# | changing the source URL of a linked picture in an Aspose.Cells workbook programmatically | refresh external image after updating its URL with Aspose.Cells for .NET | Aspose.Cells picture.Add method with URL parameter example
// Tags: Aspose.Cells linked picture insertion | modify linked picture source Aspose.Cells | refresh external image worksheet | C# add picture from URL Excel | linked image source manipulation Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a new workbook, adds a linked picture that points to an external image URL into cell B2, and saves the file as LinkedPictureDemo.xlsx. The linked picture can later be retrieved from the worksheet's Pictures collection, its LinkSource property can be set to a new URL, and Refresh can be called to update the displayed image using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a linked picture (URL) at cell B2 (row 1, column 1)
            // The picture is automatically linked to the specified URL.
            sheet.Pictures.Add(1, 1, "https://example.com/newImage.png");

            // Save the workbook
            workbook.Save("LinkedPictureDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
