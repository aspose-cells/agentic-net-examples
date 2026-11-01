// Title: How to add a custom Boolean document property named IsReviewed to an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: add a custom Boolean property to an existing Excel file with Aspose.Cells in C# | set custom document property IsReviewed to true using Aspose.Cells .NET | Aspose.Cells example for adding custom metadata to a workbook | C# code to create a Boolean custom property in an Excel workbook with Aspose.Cells | modify Excel workbook metadata by adding a custom property via Aspose.Cells
// Tags: Aspose.Cells set custom document property | Excel workbook Boolean metadata C# | custom workbook property Aspose.Cells .NET | modify Excel file custom properties | Aspose.Cells custom metadata example

using Aspose.Cells;
using System;

// The sample loads an existing Excel workbook, adds a custom Boolean document property called "IsReviewed" with a value of true, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom Boolean property named "IsReviewed" with value true
        workbook.CustomDocumentProperties.Add("IsReviewed", true);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
