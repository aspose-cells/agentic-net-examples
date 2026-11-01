// Title: Embedding a PDF as an OLE object in cell H4 with custom height and width using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# embed PDF as OLE object at cell H4 | How to set height and width for an OLE object in an Excel worksheet using Aspose.Cells | Add PDF byte array as OLE object to a specific cell with Aspose.Cells .NET | C# code to place OLE object from PDF into Excel and define its dimensions
// Tags: add oleobject pdf aspnet cells | set oleobject dimensions worksheet c# | embed pdf as oleobject aspnet | oleobjects.add method pdf bytearray | place oleobject at specific cell aspnet

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, reads a PDF file into a byte array, and uses the OleObjects.Add method to embed the PDF as an OLE object positioned at cell H4 (row 3, column 7) with a height of 200 points and a width of 300 points, then saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Path to the PDF file that will be embedded as an OLE object
            string pdfPath = "sample.pdf";

            // Verify that the PDF file exists to avoid FileNotFoundException
            if (!File.Exists(pdfPath))
            {
                Console.WriteLine($"The file \"{pdfPath}\" was not found.");
                return;
            }

            // Read the PDF file into a byte array (required by OleObjects.Add)
            byte[] pdfData = File.ReadAllBytes(pdfPath);

            // Cell H4 corresponds to row index 3 (zero‑based) and column index 7 (zero‑based)
            int upperLeftRow = 3;      // H4 row
            int upperLeftColumn = 7;   // H column

            // Desired size of the OLE object (in points)
            int height = 200; // height of the object
            int width = 300;  // width of the object

            // Add the PDF as an OLE object to the worksheet at the specified cell and size
            sheet.OleObjects.Add(upperLeftRow, upperLeftColumn, height, width, pdfData);

            // Save the workbook to a file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
