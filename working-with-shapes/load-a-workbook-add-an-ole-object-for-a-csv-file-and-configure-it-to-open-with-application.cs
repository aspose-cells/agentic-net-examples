// Title: Insert a CSV file as an OLE object at cell B2 of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Embed a CSV file into an existing workbook as an OLE object positioned at B2 with custom height and width using Aspose.Cells. | Read a CSV into a byte array and add it as an OleObject to the first worksheet of a .xlsx file with Aspose.Cells for .NET. | Create and configure an OleObject from a CSV file, set its dimensions in points, and save the updated workbook with Aspose.Cells.
// Common Searches: how to add a csv as an ole object to a specific cell using Aspose.Cells C# | Aspose.Cells embed csv file into worksheet as oleobject example | set oleobject size in points when inserting into Excel with Aspose.Cells | load workbook and insert oleobject from byte array Aspose.Cells .NET | save workbook after adding ole object from csv using Aspose.Cells
// Tags: aspocells add oleobject from csv | embed csv as oleobject in worksheet | oleobject dimensions points aspocells | load workbook insert oleobject c# | byte array oleobject aspocells | excel oleobject placement cell b2

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The program loads an existing Excel workbook, reads a CSV file into a byte array, inserts the CSV as an OLE object at cell B2 with specified height and width, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Verify that the input workbook exists
            const string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Verify that the CSV file exists
            const string csvFilePath = "data.csv";
            if (!File.Exists(csvFilePath))
                throw new FileNotFoundException($"CSV file not found: {csvFilePath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Read CSV file into a byte array (required by OleObjects.Add overload)
            byte[] oleData = File.ReadAllBytes(csvFilePath);

            // Add the OLE object (CSV) to the worksheet.
            // Parameters: start row, start column, height (points), width (points), source data as byte[]
            // Row and column are zero‑based; here we place it at cell B2 (row 1, column 1)
            int oleIndex = sheet.OleObjects.Add(
                1,                     // upper left row
                1,                     // upper left column
                200,                   // height in points
                100,                   // width in points
                oleData);              // source data

            // Retrieve the added OleObject if further manipulation is needed
            OleObject oleObject = sheet.OleObjects[oleIndex];

            // Save the modified workbook
            const string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
