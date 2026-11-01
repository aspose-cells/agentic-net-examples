// Title: Import an ArrayList of C# enum values into an Excel sheet and display enum names with a custom number format using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells ImportArrayList to write a C# enum collection to a worksheet starting at cell A1. | Create and assign a custom number format that maps each enum numeric value to its corresponding name, then apply the style to the imported range. | Save the workbook after formatting, ensuring the cells show the enum names instead of numbers.
// Common Searches: asp.net import enum list into excel cells using aspose.cells | custom number format to display enum names in excel with aspose.cells c# | how to map numeric enum values to text in an aspose.cells workbook | importarraylist example for enum values in aspose.cells .net
// Tags: ImportArrayList enum collection Aspose.Cells | custom number format enum display Excel | apply style to range Aspose.Cells | save workbook as xlsx Aspose.Cells | map numeric values to text C# Excel

using System;
using System.Collections;
using Aspose.Cells;

// The example creates a workbook, imports a C# enum ArrayList into cells A1:A5, applies a custom number format that converts the numeric enum values to their names, and saves the file as EnumDisplay.xlsx.
public class EnumImportExample
{
    // Define an enum whose values will be imported
    enum MyEnum
    {
        First = 0,
        Second = 1,
        Third = 2
    }

    public static void Main()
    {
        // Create a new workbook (lifecycle rule)
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Build an ArrayList containing enum values
        ArrayList enumValues = new ArrayList();
        enumValues.Add(MyEnum.First);
        enumValues.Add(MyEnum.Second);
        enumValues.Add(MyEnum.Third);
        enumValues.Add(MyEnum.First);
        enumValues.Add(MyEnum.Third);

        // Import the ArrayList into cells starting at row 0, column 0 (ImportArrayList rule)
        sheet.Cells.ImportArrayList(enumValues, 0, 0, false);

        // Create a style with a custom number format that maps numeric values to enum names
        Style style = workbook.CreateStyle();
        // Format: [=0]"First";[=1]"Second";[=2]"Third";General
        style.Custom = "[=0]\"First\";[=1]\"Second\";[=2]\"Third\";General";

        // Apply the style to the range where the enum values were imported (A1:A5)
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        sheet.Cells.CreateRange("A1:A5").ApplyStyle(style, flag);

        // Save the workbook (lifecycle rule)
        workbook.Save("EnumDisplay.xlsx");
    }
}
