using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Set the ScaleCrop property to true – this controls the thumbnail display mode
        workbook.BuiltInDocumentProperties.ScaleCrop = true;

        // Optional: display the current value to verify
        Console.WriteLine("ScaleCrop property value: " + workbook.BuiltInDocumentProperties.ScaleCrop);

        // Save the workbook as PDF; the ScaleCrop flag is persisted in the output file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}