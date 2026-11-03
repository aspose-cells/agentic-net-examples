// Title: Import enum values from an ArrayList into Excel and display enum names using a custom number format with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# import ArrayList of enum values into worksheet | show enum names instead of numeric values in Excel using Aspose.Cells custom format | how to map enum integers to text with ImportArrayList in Aspose.Cells | apply custom number format for enum display in Aspose.Cells .NET
// Tags: ImportArrayList enum Aspose.Cells | custom number format enum Aspose.Cells | enum display in Excel using Aspose.Cells | C# map enum to text in worksheet | Aspose.Cells format cells based on enum values

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumExample
{
    // Example enum
    enum MyEnum
    {
        Red = 1,
        Green = 2,
        Blue = 3
    }

    // The example creates a workbook, builds an ArrayList containing MyEnum values, imports the list into cells A1:A5 with ImportArrayList, defines a custom number format that maps each enum numeric value to its name, applies this style to the range, and saves the file as EnumDisplay.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare an ArrayList with enum values (as integers)
                ArrayList enumValues = new ArrayList
                {
                    MyEnum.Red,
                    MyEnum.Green,
                    MyEnum.Blue,
                    MyEnum.Red,
                    MyEnum.Blue
                };

                // Import the ArrayList into the worksheet starting at cell A1
                // Parameters: list, firstRow, firstColumn, isColumnWise
                sheet.Cells.ImportArrayList(enumValues, 0, 0, false);

                // Define a custom number format that maps enum numeric values to their names
                // [=1]"Red";[=2]"Green";[=3]"Blue";General
                string enumDisplayFormat = "[=1]\"Red\";[=2]\"Green\";[=3]\"Blue\";General";

                // Apply the custom format to the range that contains the imported values (A1:A5)
                Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, 5, 1);
                Style style = workbook.CreateStyle();
                style.Custom = enumDisplayFormat;
                range.ApplyStyle(style, new StyleFlag { NumberFormat = true });

                // Save the workbook
                string outputPath = "EnumDisplay.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
