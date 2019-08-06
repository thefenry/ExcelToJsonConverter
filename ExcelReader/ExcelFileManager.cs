using System.Collections.Generic;
using ExcelReader.Models;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ExcelFileManager
    {
        private string _filePath;
        private Application _excelApp;
        private Workbook _workBook;

        public ExcelFileManager()
        {            
            _excelApp = new Application();
        }

        public List<SheetInfo> GetFileSheets(string filePath)
        {
            _filePath = filePath;
            SetWorkBook();
            List<SheetInfo> sheets = new List<SheetInfo>();

            int sheetCount = _workBook.Sheets.Count;

            for (int sheetNumber = 1; sheetNumber < sheetCount + 1; sheetNumber++)
            {
                Worksheet workSheet = (Worksheet)_workBook.Sheets[sheetNumber];

                sheets.Add(new SheetInfo { PageNumber = sheetNumber, Name = workSheet.Name });
            }

            return sheets;
        }

        public Dictionary<string, List<Dictionary<string, string>>> GetPageContent(List<SheetInfo> sheets)
        {
            if (_workBook == null)
            {
                SetWorkBook();
            }

            Dictionary<string, List<Dictionary<string, string>>> workSheetValues = new Dictionary<string, List<Dictionary<string, string>>>();

            foreach (SheetInfo sheet in sheets)
            {
                int pageNumber = sheet.PageNumber;

                Worksheet workSheet = (Worksheet)_workBook.Sheets[pageNumber];
                if (workSheet.Name != sheet.Name)
                {
                    continue;
                }

                Range usedRange = workSheet.UsedRange;

                object[,] valueArray = (object[,])usedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                List<Dictionary<string, string>> sheetValues = new List<Dictionary<string, string>>();

                for (int i = 2; i < valueArray.GetLength(0) +1; i++)
                {
                    Dictionary<string, string> values = new Dictionary<string, string>();
                    for (int j = 1; j < valueArray.GetLength(1) + 1; j++)
                    {

                        string value = valueArray[i, j] == null ? null : valueArray[i, j].ToString();

                        values.Add(valueArray[1, j].ToString(), value);
                    }

                    sheetValues.Add(values);
                }

                workSheetValues.Add(workSheet.Name, sheetValues);
            }

            return workSheetValues;
        }

        private void SetWorkBook()
        {
            this._workBook = this._excelApp.Workbooks.Open(_filePath);
        }
    }
}
