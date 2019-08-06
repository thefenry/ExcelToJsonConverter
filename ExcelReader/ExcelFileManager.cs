using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace ExcelReader
{
    public class ExcelFileManager
    {
        private string _filePath;
        private Application _excelApp;
        private Workbook _workBook;

        public ExcelFileManager(string filePath)
        {
            _filePath = filePath;
            _excelApp = new Application();
        }

        public Dictionary<string, int> GetFileSheets()
        {
            SetWorkBook();
            Dictionary<string, int> sheetNames = new Dictionary<string, int>();
            int sheetCount = _workBook.Sheets.Count;

            for (int sheetNumber = 1; sheetNumber < sheetCount + 1; sheetNumber++)
            {
                Worksheet workSheet = (Worksheet)_workBook.Sheets[sheetNumber];
                sheetNames.Add(workSheet.Name, sheetNumber);
            }

            //using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            //{
            //    // Auto - detect format, supports:
            //    //  - Binary Excel files (2.0-2003 format; *.xls)
            //    //  - OpenXml Excel files (2007 format; *.xlsx)
            //    using (var reader = ExcelReaderFactory.CreateReader(stream))
            //    {
            //        // Choose one of either 1 or 2:

            //        // 1. Use the reader methods
            //        do
            //        {
            //            while (reader.Read())
            //            {
            //                // reader.GetDouble(0);
            //            }
            //        } while (reader.NextResult());

            //        // 2. Use the AsDataSet extension method
            //        //var result = reader.AsDataSet();

            //        // The result of each spreadsheet is in result.Tables
            //    }
            //}

            return sheetNames;
        }

        private void SetWorkBook()
        {
            this._workBook = this._excelApp.Workbooks.Open(_filePath);
        }

        public Dictionary<string, List<Dictionary<string, string>>> GetPageContent(Dictionary<string, int> sheets)
        {
            if (_workBook == null)
            {
                SetWorkBook();
            }

            //List<List<Dictionary<string, string>>> workSheetValues = new List<List<Dictionary<string, string>>>();
            Dictionary<string, List<Dictionary<string, string>>> workSheetValues = new Dictionary<string, List<Dictionary<string, string>>>();

            foreach (KeyValuePair<string, int> sheet in sheets)
            {
                int pageNumber = sheet.Value;

                Worksheet workSheet = (Worksheet)_workBook.Sheets[pageNumber];
                if (workSheet.Name != sheet.Key)
                {
                    continue;
                }

                Range usedRange = workSheet.UsedRange;

                object[,] valueArray = (object[,])usedRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);

                List<Dictionary<string, string>> sheetValues = new List<Dictionary<string, string>>();

                for (int i = 2; i < valueArray.GetLength(1); i++)
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
    }
}
