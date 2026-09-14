// Title: Combine two Excel workbooks in C# with Aspose.Cells while keeping charts and images intact
// AI Prompts: Use Aspose.Cells Workbook.Combine to merge a source .xlsx into a destination workbook and retain all embedded charts and pictures. | Generate C# code that loads two Excel files, merges them with the default combine behavior, and saves the result preserving visual objects. | Explain how the default Workbook.Combine method handles chart and image preservation during workbook merging in Aspose.Cells.
// Common Searches: how to merge two .xlsx files in C# without losing charts using Aspose.Cells | Aspose.Cells default combine method keep images when merging workbooks | C# example for combining Excel workbooks preserving embedded objects | preserve charts during Excel workbook merge Aspose.Cells
// Tags: Workbook.Combine retain visual objects | C# Aspose.Cells merge workbooks | default combine behavior Aspose.Cells | keep chart objects during Excel workbook merge | retain images in combined .xlsx

using System;
using Aspose.Cells;

namespace AsposeCellsMergeDemo
{
    // The example loads Destination.xlsx and Source.xlsx, merges them using the default Workbook.Combine method (which automatically preserves charts, images, and other embedded objects), and saves the combined file as CombinedResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Load the first workbook (destination workbook)
            Workbook destWorkbook = new Workbook("Destination.xlsx");

            // Load the second workbook (source workbook to be merged)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");

            // Combine the source workbook into the destination workbook.
            // This uses the default Workbook.Combine behavior which preserves charts, images, etc.
            destWorkbook.Combine(sourceWorkbook);

            // Save the combined workbook.
            destWorkbook.Save("CombinedResult.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbooks merged successfully. Charts and images are preserved.");
        }
    }
}
