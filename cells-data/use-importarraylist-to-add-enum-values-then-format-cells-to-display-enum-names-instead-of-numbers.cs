// Title: Import Enum Values and Display Names with ImportArrayList & Custom Number Format (Aspose.Cells C#)
// Description: Demonstrates how to import a list of ColorType enum values into an Excel worksheet using Aspose.Cells' ImportArrayList, then apply a custom number format so the cells show the enum names instead of their integer codes. The workbook is saved as EnumImportAndDisplay.xlsx.
// Keywords: Aspose.Cells ImportArrayList | C# enum to Excel | custom number format Aspose.Cells | display enum names in Excel | ColorType enum Aspose.Cells | map integer to text Excel | .NET Excel automation
// Common Searches: Aspose.Cells import enum values | show enum names instead of numbers in Excel C# | custom number format for enum display Aspose.Cells | ImportArrayList example with enums | how to map enum integers to text in Excel using Aspose
// Developer Intent: Load enum integer values into a worksheet with ImportArrayList and format the cells so the corresponding enum names are displayed.
// Use Cases: Generate Excel reports that list ColorType enum values as readable text for business users. | Populate data‑validation lists with underlying enum codes while presenting friendly names. | Create templates where enum identifiers are stored as numbers but shown as descriptive labels.
// AI Prompts: Write C# code that uses Aspose.Cells ImportArrayList to import a collection of enum values and applies a custom number format to show the enum names. | Explain how to build a custom number‑format string that maps enum integer values to their textual representations in Aspose.Cells. | Provide a step‑by‑step example of styling a range with a custom number format after importing enum values via ImportArrayList.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsEnumImportDemo
{
    // Demonstrates how to import a list of ColorType enum values into an Excel worksheet using Aspose.Cells' ImportArrayList, then apply a custom number format so the cells show the enum names instead of their integer codes. The workbook is saved as EnumImportAndDisplay.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Prepare an ArrayList with enum values (underlying integer values will be imported)
                ArrayList enumValues = new ArrayList();
                enumValues.Add(ColorType.Automatic);        // 0
                enumValues.Add(ColorType.AutomaticIndex);   // 1
                enumValues.Add(ColorType.RGB);              // 2
                enumValues.Add(ColorType.IndexedColor);     // 3
                enumValues.Add(ColorType.Theme);            // 4

                // Import the enum values horizontally starting at cell A1 (row 0, column 0)
                cells.ImportArrayList(enumValues, 0, 0, false);

                // Create a custom number format that maps each numeric value to its enum name
                // Format syntax: [=value]"DisplayText";[=value]"DisplayText";...
                string customFormat = "[=0]\"Automatic\";[=1]\"AutomaticIndex\";[=2]\"RGB\";[=3]\"IndexedColor\";[=4]\"Theme\"";

                // Create a style and assign the custom number format
                Style enumStyle = workbook.CreateStyle();
                enumStyle.Custom = customFormat;

                // Apply the style to the range that contains the imported values
                // The range spans 1 row and enumValues.Count columns
                Aspose.Cells.Range importedRange = cells.CreateRange(0, 0, 1, enumValues.Count);
                importedRange.ApplyStyle(enumStyle, new StyleFlag { NumberFormat = true });

                // Save the workbook
                workbook.Save("EnumImportAndDisplay.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
