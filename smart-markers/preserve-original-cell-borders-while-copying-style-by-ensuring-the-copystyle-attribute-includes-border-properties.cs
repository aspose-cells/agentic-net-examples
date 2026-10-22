// Title: How to copy only the border formatting from one Excel cell to another while preserving other styles using Aspose.Cells for .NET (C#)
// AI Prompts: Copy the border line style and color from cell A1 to cell B2 in an existing workbook using Aspose.Cells for .NET. | Transfer only the border properties of a source cell to a target cell without changing any other formatting in C# with Aspose.Cells. | Apply the source cell's border settings to another cell while keeping the target cell's existing style intact using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# copy border formatting between cells without affecting cell fill | preserve existing cell style while copying borders in Aspose.Cells workbook | how to clone border properties from one worksheet cell to another using Aspose.Cells for .NET | C# example to copy Excel cell borders using Aspose.Cells library
// Tags: copy cell border style Aspose.Cells C# | preserve cell formatting while cloning borders .NET | Aspose.Cells border property transfer | Excel workbook border cloning C# | style copying without affecting fill Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, reads the border line style and color of every border side from a source cell (A1), copies those border properties to a target cell (B2) while leaving other formatting unchanged, and saves the modified workbook as Output.xlsx.
class PreserveBordersExample
{
    static void Main()
    {
        try
        {
            const string inputFile = "Input.xlsx";
            const string outputFile = "Output.xlsx";

            // Ensure the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Source cell (the cell whose border style we want to copy)
            Cell sourceCell = sheet.Cells["A1"];

            // Target cell (the cell to which we want to apply the border style)
            Cell targetCell = sheet.Cells["B2"];

            // Retrieve styles from source and target cells
            Style sourceStyle = sourceCell.GetStyle();
            Style targetStyle = targetCell.GetStyle();

            // Copy border properties from source to target (LineStyle and Color)
            foreach (BorderType borderType in Enum.GetValues(typeof(BorderType)))
            {
                targetStyle.Borders[borderType].LineStyle = sourceStyle.Borders[borderType].LineStyle;
                targetStyle.Borders[borderType].Color = sourceStyle.Borders[borderType].Color;
            }

            // Apply the modified style to the target cell
            targetCell.SetStyle(targetStyle);

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
