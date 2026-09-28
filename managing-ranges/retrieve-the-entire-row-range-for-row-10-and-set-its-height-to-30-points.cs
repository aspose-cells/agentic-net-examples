// Title: Set the height of row 10 to 30 points with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to access row 10 of the first worksheet, change its Height property to 30 points, and save the workbook as an XLSX file. | Show how to obtain a Row object by index in Aspose.Cells, adjust its height, and then export the workbook.
// Common Searches: Aspose.Cells C# change height of row 10 to 30 points | How to set specific row height in an Excel file using Aspose.Cells .NET | Retrieve a row by index and adjust its Height property with Aspose.Cells | Save workbook after modifying row dimensions with Aspose.Cells C#
// Tags: Aspose.Cells set row height | C# Aspose.Cells get row object | Aspose.Cells modify row dimensions | Aspose.Cells save workbook to XLSX | Aspose.Cells row height 30 points

using Aspose.Cells;

// Creates a workbook, accesses the first worksheet, retrieves row 10 (zero‑based index 9), sets its height to 30 points, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the entire row range for row 10 (zero‑based index 9)
        Row row = sheet.Cells.Rows[9];

        // Set the row height to 30 points
        row.Height = 30;

        // Save the workbook (optional)
        workbook.Save("output.xlsx");
    }
}
