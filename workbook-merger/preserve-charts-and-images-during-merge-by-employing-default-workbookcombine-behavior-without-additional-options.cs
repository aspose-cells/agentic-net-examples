// Title: Merge Excel workbooks while preserving charts and images with Aspose.Cells Workbook.Combine (C#)
// Description: The C# sample loads Destination.xlsx and Source.xlsx, merges the source into the destination using Workbook.Combine with default settings, and saves the result as CombinedResult.xlsx. The default combine operation retains all charts, pictures, and other embedded objects.
// Keywords: Aspose.Cells Workbook.Combine | merge Excel workbooks C# | preserve charts Aspose | retain images Excel merge | .NET combine workbooks | default combine behavior | preserve embedded objects | Excel file consolidation | Aspose.Cells merge example
// Common Searches: Aspose.Cells merge workbooks keep charts | How to combine two Excel files without losing images .NET | Workbook.Combine preserve embedded objects | C# merge Excel workbooks Aspose.Cells default options | Combine multiple workbooks preserving graphics
// Developer Intent: Combine a source workbook into a destination workbook without losing any charts, pictures, or other embedded objects.
// Use Cases: Consolidate monthly financial reports that contain charts into a single workbook for executive review. | Create a master workbook from departmental files that include images, ensuring visuals remain intact after merging. | Automate the merging of a template workbook with data workbooks while retaining all visual objects for a reporting pipeline.
// AI Prompts: Generate C# code that uses Aspose.Cells Workbook.Combine to merge two Excel files and keep all charts and images. | Explain why Workbook.Combine's default settings preserve embedded objects and how to verify the result after merging. | Provide sample C# that loops through a list of workbooks, merging each into a destination workbook while retaining every chart and picture.

using System;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    // The C# sample loads Destination.xlsx and Source.xlsx, merges the source into the destination using Workbook.Combine with default settings, and saves the result as CombinedResult.xlsx. The default combine operation retains all charts, pictures, and other embedded objects.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the first workbook (destination workbook)
            Workbook destWorkbook = new Workbook("Destination.xlsx");

            // Load the second workbook (source workbook to be merged)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");

            // Combine the source workbook into the destination workbook.
            // Default Combine behavior preserves charts, images, and other objects.
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook.
            destWorkbook.Save("CombinedResult.xlsx", SaveFormat.Xlsx);
        }
    }
}
