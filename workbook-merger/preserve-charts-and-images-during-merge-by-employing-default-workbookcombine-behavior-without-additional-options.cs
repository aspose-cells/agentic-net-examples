// Title: Combine two Excel workbooks in C# while preserving charts and images using Aspose.Cells
// Common Searches: c# combine two .xlsx files without losing charts Aspose.Cells | how to merge Excel workbooks and retain images using Aspose.Cells | Aspose.Cells default Combine method keep drawing objects | merge workbooks preserving charts and pictures in C# | combine Excel files preserving embedded charts Aspose.Cells
// Tags: Aspose.Cells Combine preserve charts | merge Excel workbooks with images C# | default workbook combine drawing objects | preserve charts during workbook merge | combine .xlsx files without extra options

using System;
using Aspose.Cells;

// The example loads two Excel workbooks, uses the default Workbook.Combine method to merge the source workbook into the target workbook while retaining charts, images, and other drawing objects, and saves the merged result as a new file.
class Program
{
    static void Main()
    {
        // Load the first workbook (the target workbook)
        Workbook targetWorkbook = new Workbook("SourceWorkbook1.xlsx");

        // Load the second workbook (the workbook to be merged into the target)
        Workbook sourceWorkbook = new Workbook("SourceWorkbook2.xlsx");

        // Combine the source workbook into the target workbook.
        // Default Combine behavior preserves charts, images, and other drawing objects.
        targetWorkbook.Combine(sourceWorkbook);

        // Save the merged workbook to a new file.
        targetWorkbook.Save("MergedWorkbook.xlsx");
    }
}
