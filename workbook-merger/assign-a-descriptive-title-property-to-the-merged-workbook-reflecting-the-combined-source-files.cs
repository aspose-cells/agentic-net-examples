using System;
using Aspose.Cells;

// Author: Aspose.Cells .NET example - combines workbooks and sets a descriptive title
class Program
{
    static void Main()
    {
        // Create first source workbook
        Workbook source1 = new Workbook();
        source1.Worksheets[0].Cells["A1"].PutValue("Data from Source 1");

        // Create second source workbook
        Workbook source2 = new Workbook();
        source2.Worksheets[0].Cells["A1"].PutValue("Data from Source 2");

        // Combine the second workbook into the first one
        source1.Combine(source2);

        // Assign a descriptive title reflecting the combined source files
        source1.BuiltInDocumentProperties.Title = "Combined Workbook: Source1 + Source2";

        // Save the merged workbook with the title set
        source1.Save("CombinedWorkbookWithTitle.xlsx", SaveFormat.Xlsx);
    }
}