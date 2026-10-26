// Title: Generate an Excel workbook with twelve month‑named worksheets and fill each sheet with sample data using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create workbook with worksheets for each month | populate multiple Excel sheets with the same 2D array using Aspose.Cells | apply bold gray header style to worksheet rows in Aspose.Cells .NET | auto‑fit columns on every worksheet programmatically with Aspose.Cells | save Aspose.Cells workbook as .xlsx file in C#
// Tags: create month worksheets Aspose.Cells | fill worksheets from 2d string array C# | header row bold gray style Aspose.Cells | auto fit columns each sheet Aspose.Cells | save workbook as xlsx C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, adds twelve worksheets named January through December, populates each with a predefined 2‑D string array of sample items, applies a bold light‑gray header style, auto‑fits all columns, and saves the file as MonthsData.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define month names
        string[] months = new string[]
        {
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        };

        // Sample data to populate each worksheet
        string[,] sampleData = new string[,]
        {
            { "Item", "Quantity", "Price" },
            { "Apples", "120", "0.50" },
            { "Bananas", "80", "0.30" },
            { "Oranges", "150", "0.45" }
        };

        // Loop through each month, create a worksheet, and fill it with sample data
        foreach (string month in months)
        {
            // Add a new worksheet with the month name
            Worksheet sheet = workbook.Worksheets.Add(month);

            // Populate the worksheet with the sample data
            for (int row = 0; row < sampleData.GetLength(0); row++)
            {
                for (int col = 0; col < sampleData.GetLength(1); col++)
                {
                    sheet.Cells[row, col].PutValue(sampleData[row, col]);
                }
            }

            // Optionally, apply a simple style to the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the first row (header)
            for (int col = 0; col < sampleData.GetLength(1); col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Auto-fit columns for better appearance
            sheet.AutoFitColumns();
        }

        // Save the workbook to a file
        workbook.Save("MonthsData.xlsx");
    }
}
