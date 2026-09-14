// Title: Convert numeric strings to true numbers only on the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel file with Aspose.Cells, call Cells.ConvertStringToNumericValue on the first worksheet, and save the workbook without altering other sheets. | Use the ConvertStringToNumericValue method on Worksheet[0] to transform string representations of numbers into numeric cells while keeping remaining worksheets intact. | Demonstrate how to apply numeric string conversion to a single sheet in a multi‑sheet workbook and write the result to a new file with Aspose.Cells C#.
// Common Searches: Aspose.Cells C# convert string numbers to numeric on first sheet only | How to use ConvertStringToNumericValue for a specific worksheet in a workbook | Preserve other worksheets while converting numeric strings in Aspose.Cells | C# example converting numeric strings in sheet index 0 with Aspose.Cells | Apply numeric conversion to one worksheet and save workbook Aspose.Cells
// Tags: Aspose.Cells ConvertStringToNumericValue first worksheet | numeric string conversion C# Aspose.Cells | preserve other sheets during numeric conversion Aspose | load and save workbook after numeric conversion Aspose.Cells | single sheet numeric conversion Excel Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumericDemo
{
    // Loads an Excel workbook, applies Cells.ConvertStringToNumericValue to the first worksheet to turn numeric strings into real numbers, leaves all other worksheets unchanged, and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet (index 0)
            Worksheet firstSheet = workbook.Worksheets[0];

            // Convert all string values that can be interpreted as numbers to numeric values
            // This operation is performed only on the first worksheet
            firstSheet.Cells.ConvertStringToNumericValue();

            // Save the workbook to a new file (replace with desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
