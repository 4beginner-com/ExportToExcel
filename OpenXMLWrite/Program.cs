using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
internal class Program
{
    private static void Main(string[] args)
    {
        string filePath = @"C:\Works\OpenXMLWrite\";
        // Header data setup
        Dictionary<string, string> headerData = new Dictionary<string, string>
        {
            { "CountryId", "Country ID" },
            { "CountryCode", "Country Code" },
            { "CountryName", "Country Name" }
        };

        // Sample data setup
        List<Dictionary<string, object>> data = new List<Dictionary<string, object>>
        {
            new Dictionary<string, object>
            {
                { "CountryId", 1 },
                { "CountryCode", "1"},
                { "CountryName", "United States of America" }
            },
            new Dictionary<string, object>
            {
                { "CountryId", 2 },
                { "CountryCode", "91" },
                { "CountryName", "Canada" }
            },
            new Dictionary<string, object>
            {
                { "CountryId", 3 },
                { "CountryCode", "44" },
                { "CountryName", "United Kingdom" }
            },
            // Add more countries as needed
        };
        var excelData = WriteToExcel("Sample.xlsx", filePath, data, headerData);
        var memStream = SetHeaderRowStyle(excelData, "00FF00");
        WriteMemoryStreamToFile(memStream, filePath + "output.xlsx");
    }

    public static MemoryStream WriteToExcel(string fileName, string filePath, List<Dictionary<string, object>> data, Dictionary<string, string> headerData)
    {
        string fullPath = Path.Combine(filePath, fileName);
        if (!File.Exists(fullPath))
        {
            return null;
        }

        MemoryStream memoryStream = new MemoryStream();
        using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
        {
            fs.CopyTo(memoryStream);
        }

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, true))
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            if (sheet == null)
            {
                return null; // No sheet found, or you could create one if necessary
            }

            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Writing header
            if (headerData.Count > 0)
            {
                Row headerRow = new Row { RowIndex = 1 };
                sheetData.AppendChild(headerRow);
                int columnIndex = 0;
                foreach (var header in headerData)
                {
                    Cell cell = CreateTextCell(ColumnLetter(columnIndex++), 1, header.Value);
                    headerRow.AppendChild(cell);
                }
            }

            // Writing data rows
            for (int i = 0; i < data.Count; i++)
            {
                Row dataRow = new Row { RowIndex = (uint)(i + 2) }; // Adjusting row index to account for header
                sheetData.AppendChild(dataRow);

                int columnIndex = 0;
                foreach (var valuePair in data[i])
                {
                    if (headerData.ContainsKey(valuePair.Key))
                    {
                        string cellValue = valuePair.Value?.ToString() ?? string.Empty;
                        Cell cell = CreateTextCell(ColumnLetter(columnIndex++), (uint)i + 2, cellValue);
                        dataRow.AppendChild(cell);
                    }
                }
            }

            worksheetPart.Worksheet.Save();
        }

        // Resetting memory stream position for reading
        memoryStream.Position = 0;
        return memoryStream;
    }

    private static Cell CreateTextCell(string header, uint index, string text)
    {
        return new Cell
        {
            CellReference = header + index,
            DataType = CellValues.String,
            CellValue = new CellValue(text)
        };
    }

    // Helper to convert column index to letter (simplified, no support for AA, AB, etc.)
    private static string ColumnLetter(int intCol)
    {
        int intFirstLetter = ((intCol) / 676) + 64;
        int intSecondLetter = ((intCol % 676) / 26) + 64;
        int intThirdLetter = (intCol % 26) + 65;

        char firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
        char secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
        char thirdLetter = (char)intThirdLetter;

        return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
    }

    public static void WriteMemoryStreamToFile(MemoryStream memoryStream, string outputPath)
    {
        // Ensure the memory stream is at the beginning
        memoryStream.Seek(0, SeekOrigin.Begin);

        // Using 'using' statement to ensure the fileStream is disposed properly
        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            memoryStream.WriteTo(fileStream);
        }
    }

    //Header to Yellow
    public static MemoryStream SetHeaderRowStyle(MemoryStream inputStream, string colorHex)
    {
        // Create a new MemoryStream for output
        MemoryStream outputStream = new MemoryStream();

        // Copy the input stream to the output stream to work on a fresh copy
        inputStream.Seek(0, SeekOrigin.Begin); // Ensure we're copying from the beginning
        inputStream.CopyTo(outputStream);

        // Ensure the outputStream is positioned at the beginning for reading
        outputStream.Seek(0, SeekOrigin.Begin);

        using (SpreadsheetDocument document = SpreadsheetDocument.Open(outputStream, true))
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            // Ensure Stylesheet is initialized
            WorkbookStylesPart stylesPart = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (stylesPart == null)
            {
                stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(new Fills(), new Borders(), new Fonts(), new CellFormats());
            }

            Stylesheet stylesheet = stylesPart.Stylesheet;

            // Ensure Fills is initialized and add a fill with the specified color
            if (stylesheet.Fills == null) stylesheet.Fills = new Fills();
            Fill customFill = new Fill(
                new PatternFill(
                    new ForegroundColor { Rgb = colorHex } // Use the colorHex parameter for color
                )
                { PatternType = PatternValues.Solid }
            );
            stylesheet.Fills.AppendChild(customFill);
            stylesheet.Fills.Count++;

            // Ensure CellFormats is initialized and add a new format using the custom fill
            if (stylesheet.CellFormats == null) stylesheet.CellFormats = new CellFormats();
            CellFormat cellFormatUsingCustomFill = new CellFormat { FillId = stylesheet.Fills.Count - 1, ApplyFill = true };
            stylesheet.CellFormats.AppendChild(cellFormatUsingCustomFill);
            stylesheet.CellFormats.Count++;

            // Apply the new style to header row cells
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            Row headerRow = sheetData.Elements<Row>().FirstOrDefault();
            if (headerRow != null)
            {
                foreach (Cell cell in headerRow.Elements<Cell>())
                {
                    cell.StyleIndex = stylesheet.CellFormats.Count - 1;
                }
            }

            // Save changes
            stylesheet.Save();
            worksheetPart.Worksheet.Save();
        }

        // The outputStream now contains the modified document
        // Reset the position of the outputStream to allow for reading from the beginning
        outputStream.Seek(0, SeekOrigin.Begin);

        return outputStream;
    }

}