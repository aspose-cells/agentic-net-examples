using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnableThumbnailScalingDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable scaling of the document thumbnail
            workbook.BuiltInDocumentProperties.ScaleCrop = true;

            // Optional: display the current setting
            Console.WriteLine("ScaleCrop property value: " + workbook.BuiltInDocumentProperties.ScaleCrop);

            // Save the workbook
            workbook.Save("ThumbnailScaled.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            EnableThumbnailScalingDemo.Run();
        }
    }
}