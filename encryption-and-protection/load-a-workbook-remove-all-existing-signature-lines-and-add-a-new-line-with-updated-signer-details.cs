using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string inputPath = "SignedWorkbook.xlsx";
        string outputPath = "UpdatedWorkbook.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputPath);

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            ShapeCollection shapes = sheet.Shapes;

            // Identify indices of shapes that contain a signature line
            List<int> indicesToRemove = new List<int>();
            for (int i = 0; i < shapes.Count; i++)
            {
                // A signature line is stored in a Picture object's SignatureLine property
                if (shapes[i] is Picture picture && picture.SignatureLine != null)
                {
                    indicesToRemove.Add(i);
                }
            }

            // Remove the identified signature line shapes (reverse order to keep indices valid)
            for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                shapes.RemoveAt(indicesToRemove[i]);
            }

            // Create a new signature line with updated signer details
            SignatureLine newSignature = new SignatureLine
            {
                Signer = "Alice Johnson",
                Title = "Chief Financial Officer",
                Email = "alice.johnson@example.com",
                IsLine = true,
                AllowComments = true,
                ShowSignedDate = true,
                Instructions = "Please sign to approve the financial report."
            };

            // Add the new signature line to the worksheet at row 2, column 2 (zero‑based indices)
            shapes.AddSignatureLine(1, 1, newSignature);
        }

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}