using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Define source and destination file paths
        string sourcePath = "source.xlsx";
        string destinationPath = "destination.xlsx";

        // Load the source workbook (contains timelines, formatting, and cell relationships)
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Create an empty workbook that will receive the copied content
        Workbook destinationWorkbook = new Workbook();

        // Copy all worksheets, styles, charts, timelines, and other objects from the source
        // The Copy method preserves formatting and cell relationships.
        destinationWorkbook.Copy(sourceWorkbook);

        // Save the new workbook with the copied timeline data
        destinationWorkbook.Save(destinationPath);
    }
}