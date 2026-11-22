// Title: Load an unencrypted Excel workbook and confirm its file format using Aspose.Cells for .NET
// Common Searches: how to open an unencrypted .xlsx file with Aspose.Cells in C# and read its format | Aspose.Cells C# check workbook FileFormat after loading | determine Excel file type using Aspose.Cells Workbook.FileFormat property | C# console app verify Excel workbook format with Aspose.Cells | load Excel workbook without password and get format Aspose.Cells .NET
// Tags: Aspose.Cells load unencrypted workbook | Workbook.FileFormat property C# | detect Excel file format with Aspose.Cells | verify workbook format after loading | console output workbook format Aspose.Cells

using System;
using Aspose.Cells;

// The example demonstrates loading an unencrypted Excel file (input.xlsx) with Aspose.Cells' Workbook class, retrieving the Workbook.FileFormat property to identify the workbook's format, and writing a confirmation message to the console indicating either the detected format or that the format could not be determined.
class Program
{
    static void Main()
    {
        // Path to the workbook file (unencrypted)
        string filePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Verify the workbook format
        FileFormatType format = workbook.FileFormat;

        // Log confirmation to the console
        if (format != FileFormatType.Unknown)
        {
            Console.WriteLine($"Workbook loaded successfully. Detected format: {format}");
        }
        else
        {
            Console.WriteLine("Unable to determine the workbook format.");
        }
    }
}
