using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a linked picture with an initial source URL
        string initialUrl = "https://example.com/initial.jpg";
        Picture picture = worksheet.Shapes.AddLinkedPicture(1, 1, 100, 100, initialUrl);

        // Output initial state
        Console.WriteLine("Is linked: " + picture.IsLink);
        Console.WriteLine("Initial SourceFullName: " + picture.SourceFullName);

        // Change the picture's source to a new URL
        string newUrl = "https://example.com/updated.jpg";
        picture.SourceFullName = newUrl;

        // Output updated state
        Console.WriteLine("Updated SourceFullName: " + picture.SourceFullName);

        // Save the workbook
        workbook.Save("LinkedPictureUpdated.xlsx");
    }
}