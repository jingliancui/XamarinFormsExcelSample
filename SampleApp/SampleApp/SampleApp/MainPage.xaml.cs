using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Essentials;
using Xamarin.Forms;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace SampleApp
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
            
        }

        private const string fileName = "xamarinlibrary.xlsx";
        
        protected override void OnAppearing()
        {
            var dir = FileSystem.AppDataDirectory;
            var filepath = $"{dir}/{fileName}";
            if (System.IO.File.Exists(filepath))
            {
                System.IO.File.Delete(filepath);
            }            
        }

        public void CreateSpreadsheetWorkbook(string filepath)
        {
            var dir = FileSystem.AppDataDirectory;
            filepath = $"{dir}/{filepath}";

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            // Get the sheetData cell table.
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Add a row to the cell table.
            Row row;
            row = new Row() { RowIndex = 1 };
            sheetData.Append(row);

            // In the new row, find the column location to insert a cell in A1.  
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, "A1", true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            // Add the cell to the cell table at A1.
            Cell newCell = new Cell() { CellReference = "A1" };
            row.InsertBefore(newCell, refCell);

            // Set the cell value to be a numeric value of 100.
            newCell.CellValue = new CellValue("100");
            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);

            // Close the document.
            spreadsheetDocument.Close();
        }

        private void CreateXlsxBtn_Clicked(object sender, EventArgs e)
        {
            CreateSpreadsheetWorkbook(fileName);
            CreateStatusLabel.Text = "Status:Create Finished";
        }

        private void OpenXlsxBtn_Clicked(object sender, EventArgs e)
        {
            var dir = FileSystem.AppDataDirectory;
            var filepath = $"{dir}/{fileName}";
            var doc = SpreadsheetDocument.Open(filepath, true);
            var sheetData = doc.WorkbookPart.WorksheetParts.FirstOrDefault().Worksheet.GetFirstChild<SheetData>();
            var firstRow = sheetData.ChildElements.FirstOrDefault() as Row;
            var firstCell = firstRow.ChildElements.FirstOrDefault() as Cell;
            var cellVal = firstCell.CellValue.Text;
            FirstValLabel.Text = $"First value:{cellVal}";
        }

        //https://stackoverflow.com/questions/23102010/open-xml-reading-from-excel-file
        private void ReadResourceFile_Clicked(object sender, EventArgs e)
        {
            var resourceID = "SampleApp.Files.XamarinLibraryExcel.xlsx";
            var assembly = Assembly.Load("SampleApp");

            using (Stream stream = assembly.GetManifestResourceStream(resourceID))
            {
                var doc = SpreadsheetDocument.Open(stream, false);

                var table = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable;
                var item = table.FirstChild as SharedStringItem;
                
                ResourceValLabel.Text = $"Resource First Value:{item.Text.Text}";
            }
        }
    }
}
