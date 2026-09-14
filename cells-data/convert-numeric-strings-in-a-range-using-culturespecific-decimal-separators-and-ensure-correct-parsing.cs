// Title: Convert culture‑specific numeric strings to numbers in an Aspose.Cells workbook (C#)
// AI Prompts: Configure a workbook's NumberDecimalSeparator and NumberGroupSeparator, then invoke ConvertStringToNumericValue to turn locale‑formatted strings into numeric cells. | Write European‑style numbers such as "1,23" or "12.345,67" into a worksheet range and automatically parse them to double values using Aspose.Cells. | Demonstrate how to keep non‑numeric text unchanged while converting only parsable strings to numeric values in an Aspose.Cells sheet.
// Common Searches: Aspose.Cells C# convert comma decimal strings to numeric values in Excel | How to set custom decimal and group separators for a workbook in Aspose.Cells | ConvertStringToNumericValue usage with European number formats in .NET | Parsing locale‑specific numeric strings in Aspose.Cells worksheet | Preserve text cells when converting strings to numbers with Aspose.Cells
// Tags: custom decimal separator conversion Aspose.Cells | culture‑aware numeric parsing in Excel workbook | string to numeric conversion with culture settings | preserve text cells during numeric conversion | Aspose.Cells number format handling C#

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a Workbook, sets the decimal separator to ',' and the group separator to '.', writes several strings using those separators into cells A1‑A4, calls ConvertStringToNumericValue to convert parsable strings into numeric values, prints the results, and saves the workbook as ConvertedNumbers.xlsx.
    public class ConvertNumericStringsWithCultureDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Set culture‑specific separators (comma as decimal, dot as group)
                wb.Settings.NumberDecimalSeparator = ',';
                wb.Settings.NumberGroupSeparator = '.';

                // Get the cells collection of the first worksheet
                Cells cells = wb.Worksheets[0].Cells;

                // Populate cells with numeric strings that use the specified separators
                cells["A1"].PutValue("1,23");          // 1.23
                cells["A2"].PutValue("4,567");         // 4.567
                cells["A3"].PutValue("12.345,67");     // 12,345.67
                cells["A4"].PutValue("Not a number"); // stays as string

                // Convert strings that can be parsed to numeric values
                cells.ConvertStringToNumericValue();

                // Output the converted values
                Console.WriteLine("A1 numeric: " + cells["A1"].DoubleValue);
                Console.WriteLine("A2 numeric: " + cells["A2"].DoubleValue);
                Console.WriteLine("A3 numeric: " + cells["A3"].DoubleValue);
                Console.WriteLine("A4 remains string: " + cells["A4"].StringValue);

                // Save the workbook (optional)
                wb.Save("ConvertedNumbers.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertNumericStringsWithCultureDemo.Run();
        }
    }
}
