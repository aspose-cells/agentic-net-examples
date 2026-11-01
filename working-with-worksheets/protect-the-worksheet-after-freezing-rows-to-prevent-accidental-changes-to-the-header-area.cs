// Title: How to freeze the top row and protect the header area in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# freeze first row and lock header cells | protect worksheet while keeping frozen panes with Aspose.Cells .NET | lock header row after using FreezePanes in Aspose.Cells C# example | prevent editing of header area in Excel using Aspose.Cells protection | apply locked style to header and unlock data rows Aspose.Cells C#
// Tags: FreezePanes top row Aspose.Cells | lock header cells with IsLocked style Aspose.Cells | worksheet Protect method after FreezePanes Aspose.Cells | unlock data cells while protecting header Aspose.Cells | save protected Excel workbook C# Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, adds a header row and sample data, freezes the first row with FreezePanes, applies a locked style to the header cells and an unlocked style to the data cells, protects the worksheet to prevent changes to the locked header, and saves the result as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate header row (row 0) and some data rows
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["C1"].PutValue("Score");

            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["C2"].PutValue(85);

            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");
            cells["C3"].PutValue(92);

            // Freeze the first row so it stays visible while scrolling
            // FreezePanes(row, column, totalRows, totalColumns)
            // Freeze 1 row (row index 1) and 0 columns
            sheet.FreezePanes(1, 0, 1, 0);

            // Lock the header cells (row 0) and unlock the rest
            Style headerStyle = workbook.CreateStyle();
            headerStyle.IsLocked = true; // Ensure header is locked

            Style unlockedStyle = workbook.CreateStyle();
            unlockedStyle.IsLocked = false; // Allow editing of data cells

            // Apply locked style to header row
            for (int col = 0; col < 3; col++)
            {
                cells[0, col].SetStyle(headerStyle);
            }

            // Apply unlocked style to data rows
            for (int row = 1; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].SetStyle(unlockedStyle);
                }
            }

            // Protect the worksheet to prevent changes to locked cells (header area)
            sheet.Protect(ProtectionType.All);

            // Save the workbook to a file
            workbook.Save("ProtectedWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
