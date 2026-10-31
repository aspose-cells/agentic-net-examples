// Title: Add a cell comment and save the workbook with LightCellsDataProvider while preserving annotations using Aspose.Cells for .NET
// Common Searches: asp.net add comment to a cell and export workbook with LightCellsDataProvider | preserve cell notes when saving a workbook using Aspose.Cells C# | how to write cell annotations with LightCellsDataProvider in Aspose.Cells | save workbook with comments to XLSX using Aspose.Cells for .NET | C# example of adding a note to a cell and exporting with LightCellsDataProvider
// Tags: add cell comment Aspose.Cells C# | save workbook with comments LightCellsDataProvider | preserve annotations XLSX Aspose.Cells | write cell note using Aspose.Cells API | export workbook with notes .NET

using System;
using Aspose.Cells;

namespace AsposeCellsLightCellsComments
{
    // The example creates a workbook, writes "Hello World" to cell A1, adds a comment "This is a user annotation" to that cell, and then saves the file as OutputWithComments.xlsx using Aspose.Cells, ensuring the comment is retained in the exported XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Add a value to cell A1
                ws.Cells["A1"].PutValue("Hello World");

                // Add a comment to cell A1
                Comment comment = ws.Comments[0, 0];
                comment.Note = "This is a user annotation";

                // Save the workbook with comments
                string outputPath = "OutputWithComments.xlsx";
                wb.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
