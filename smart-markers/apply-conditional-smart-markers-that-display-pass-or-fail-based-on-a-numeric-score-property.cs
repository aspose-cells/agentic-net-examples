using System;
using System.Collections.Generic;
using Aspose.Cells;

// Data class containing the numeric score
public class ScoreData
{
    public int Score { get; set; }
}

public class ConditionalSmartMarkerDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a conditional smart marker.
        // The expression evaluates the Score property:
        // if Score >= 60 => "Pass", otherwise => "Fail"
        sheet.Cells["A1"].PutValue("&=$Score>=60?\"Pass\":\"Fail\"");

        // Prepare a list of data objects to be bound to the smart marker
        List<ScoreData> scores = new List<ScoreData>
        {
            new ScoreData { Score = 85 }, // Should display "Pass"
            new ScoreData { Score = 55 }  // Should display "Fail"
        };

        // Set up the WorkbookDesigner, assign the workbook and the data source
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = workbook
        };
        designer.SetDataSource("Data", scores);

        // Process the smart markers (default LineByLine = true)
        designer.Process();

        // Save the populated workbook
        workbook.Save("ConditionalSmartMarkerOutput.xlsx");
    }
}