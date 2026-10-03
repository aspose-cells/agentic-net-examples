// Title: Merge Excel workbooks while preserving charts and images using Aspose.Cells Workbook.Combine in C#
// AI Prompts: Write C# code that loads multiple .xlsx files and merges them into a single workbook with Workbook.Combine, ensuring all charts and pictures stay intact. | Show how to combine several source workbooks into an empty destination workbook using Aspose.Cells so that embedded visual objects are not lost. | Provide a step‑by‑step example of merging Excel files with Aspose.Cells where the default Combine method retains charts, images, and other objects.
// Common Searches: aspnet merge multiple excel files keep charts and images | c# Aspose.Cells combine workbooks preserve embedded objects | how to retain charts when merging .xlsx files using Aspose.Cells | default Workbook.Combine behavior for images and charts in C#
// Tags: Aspose.Cells combine operation for .xlsx files | preserve charts during workbook merge C# | retain embedded pictures Aspose.Cells merge | default merge behavior Aspose.Cells workbook | C# merge multiple Excel workbooks with visual objects

using System;
using Aspose.Cells;

// // This program creates an empty destination workbook, loads two source workbooks that contain charts and images, merges them using the default Workbook.Combine method (which retains all visual objects), and saves the combined result as MergedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create an empty destination workbook
        Workbook destination = new Workbook();

        // Load source workbooks that contain charts and images
        Workbook source1 = new Workbook("SourceWorkbook1.xlsx");
        Workbook source2 = new Workbook("SourceWorkbook2.xlsx");

        // Combine the first source workbook into the destination.
        // The default Combine method preserves charts, images, and other objects.
        destination.Combine(source1);

        // Combine the second source workbook into the destination.
        destination.Combine(source2);

        // Save the merged workbook. Charts and images are retained.
        destination.Save("MergedWorkbook.xlsx");
    }
}
