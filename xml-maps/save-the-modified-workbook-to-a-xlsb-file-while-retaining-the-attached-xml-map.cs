// Title: Save a modified Excel workbook as .xlsb while preserving an attached XML map using Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file that contains an XML map, change a cell value, and save it as a binary .xlsb workbook without losing the map using Aspose.Cells in C#. | Convert a workbook with an attached XML map to .xlsb format while keeping the map intact, leveraging the Aspose.Cells Save method in C#.
// Common Searches: Aspose.Cells C# save workbook as xlsb preserving XML map | How to retain attached XML map when converting xlsx to xlsb with Aspose.Cells | C# code to modify a cell and export to binary Excel format without losing XML mapping | Saving Excel binary file (.xlsb) with existing XML map using Aspose.Cells .NET | Preserve XML map during workbook conversion to .xlsb in C#
// Tags: Aspose.Cells save workbook as xlsb with XML map | C# modify cell retain XML map | binary Excel export preserving XML mapping | convert xlsx to xlsb Aspose.Cells | XML map retention during workbook save

using System;
using Aspose.Cells;

// The example loads a workbook that already has an XML map, updates cell A1, and saves the file as a binary .xlsb workbook. Aspose.Cells automatically retains the attached XML map during the save operation.
class Program
{
    static void Main()
    {
        // Load the workbook that already contains an attached XML map
        Workbook workbook = new Workbook("input.xlsx");

        // Example modification: change a cell value
        workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // Save the workbook as .xlsb; the XML map is retained automatically
        workbook.Save("output.xlsb", SaveFormat.Xlsb);
    }
}
