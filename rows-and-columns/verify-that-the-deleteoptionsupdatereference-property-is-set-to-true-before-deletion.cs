// Title: How to verify DeleteOptions.UpdateReference is true before deleting a column with Aspose.Cells for .NET
// AI Prompts: Write C# code that checks DeleteOptions.UpdateReference is enabled before calling Worksheet.Cells.DeleteColumns in Aspose.Cells. | Show the steps to configure a DeleteOptions object with UpdateReference = true and then delete a column while preserving formulas. | Provide a sample that validates the DeleteOptions settings at runtime and performs column deletion only when UpdateReference is set. | Demonstrate how to programmatically ensure formula references are updated when removing a column using Aspose.Cells DeleteOptions.
// Common Searches: Aspose.Cells DeleteOptions.UpdateReference default value and how to set it | C# delete a worksheet column and keep formulas updated with Aspose.Cells | how to check DeleteOptions properties before calling DeleteColumns in Aspose.Cells | preserve cell references when removing a column using Aspose.Cells DeleteOptions
// Tags: Aspose.Cells DeleteOptions.UpdateReference | C# column deletion preserving formulas Aspose.Cells | DeleteOptions configuration for column removal | Aspose.Cells column delete with reference update | verify DeleteOptions before worksheet column deletion

using System;
using Aspose.Cells;

// Illustrates creating a workbook, setting DeleteOptions.UpdateReference to true, deleting the first column while keeping formula references intact, and saving the workbook using Aspose.Cells for .NET.
public class DeleteOptionsVerificationDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data and formulas that reference column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].Formula = "=A1*2";
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["B2"].Formula = "=A2*2";

            // Create DeleteOptions instance and ensure UpdateReference is true
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete one column starting at index 0 using the configured options
            sheet.Cells.DeleteColumns(0, 1, options);

            // Save the modified workbook
            workbook.Save("DeleteOptionsVerificationDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        DeleteOptionsVerificationDemo.Run();
    }
}
