// Title: How to format cells in Aspose.Cells .NET to display numbers in scientific notation with two decimal places
// AI Prompts: Generate C# code that creates a workbook, defines a style with the custom format "0.00E+00", and applies it to a specified range using a StyleFlag. | Write a method that receives a Worksheet and a cell address range, then sets those cells to show values in scientific notation with exactly two decimal places.
// Common Searches: Aspose.Cells C# set custom number format 0.00E+00 for a range | How to display numbers in scientific notation with two decimal places in Excel using Aspose.Cells | Apply number format to specific cells without affecting other styles in Aspose.Cells | Saving workbook after applying scientific notation formatting with Aspose.Cells .NET
// Tags: custom scientific notation format Aspose.Cells | apply custom format to range Aspose.Cells | use StyleFlag for number format Aspose.Cells | save Excel with scientific notation Aspose.Cells | C# Aspose.Cells number format customization

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, inserting numeric values, defining a style with the custom format "0.00E+00" to show two decimal places in scientific notation, applying the style to cells A1:A3 via a StyleFlag, and saving the file as ScientificNotation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert sample numeric values
            sheet.Cells["A1"].PutValue(12345);
            sheet.Cells["A2"].PutValue(0.00123);
            sheet.Cells["A3"].PutValue(-987654321);

            // Create a style with a custom number format for scientific notation
            // "0.00E+00" shows two decimal places in scientific notation
            Style sciStyle = workbook.CreateStyle();
            sciStyle.Custom = "0.00E+00";

            // Apply the style to the desired range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:A3");
            // Ensure only the number format is applied
            StyleFlag flag = new StyleFlag();
            flag.NumberFormat = true;
            range.ApplyStyle(sciStyle, flag);

            // Save the workbook (lifecycle save rule)
            workbook.Save("ScientificNotation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
