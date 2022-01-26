using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace FundIndexMatch
{
    class ExcelData
    {
        public ExcelData(string filePath, int dataSheetCount)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ep = new ExcelPackage(new FileInfo(filePath));

            Errors = new List<ExcelDataError>();
            this.dataSheetCount = dataSheetCount;
            rates = LoadRates();
            indexData = LoadIndexData();
        }

        public void WriteMatches()
        {
            try
            {
                foreach (var sheetName in indexData.Keys)
                {
                    var ws = ep.Workbook.Worksheets[sheetName];
                    MatchSheetData(ws);
                }
            }
            finally
            {
                WriteErrors();
                ep.Save();
            }
        }

        readonly ExcelPackage ep;
        readonly Dictionary<DateTime, decimal> rates;
        readonly List<ExcelDataError> Errors;
        readonly Dictionary<string, Dictionary<DateTime, decimal>> indexData;
        readonly int dataSheetCount;

        private Dictionary<DateTime, decimal> LoadRates()
        {
            var ws = ep.Workbook.Worksheets["Kurz dolaru"];
            return MapData(ws, 1);
        }

        private Dictionary<DateTime, decimal> GetData(ExcelWorksheet ews)
        {
            return MapData(ews, 3);
        }

        private Dictionary<DateTime, decimal> MapData(ExcelWorksheet ews, int headerPos)
        {
            var ret = new Dictionary<DateTime, decimal>();

            int colDate;
            try
            {
                colDate = GetHeaderColumn(ews, headerPos);
            }
            catch (IndexOutOfRangeException)
            {
                Errors.Add(new ExcelDataError(ews.Name, "", ExcelDataErrorType.MissingHeader));
                return ret;
            }

            int colValue = colDate + 1;

            int firstRow = ews.Dimension.Start.Row + 1;
            int lastRow = ews.Dimension.End.Row;

            for (var row = firstRow; row <= lastRow; ++row)
            {
                var cellDate = ews.Cells[row, colDate];
                var cellValue = ews.Cells[row, colValue];
                if (string.IsNullOrEmpty(cellDate.Text))
                    continue;

                if (!DateTime.TryParse(cellDate.Text, out DateTime date))
                {
                    Errors.Add(new ExcelDataError(ews.Name, cellDate.Address, ExcelDataErrorType.ExpectedDate));
                    continue;
                }

                try
                {
                    ret[date] = Convert.ToDecimal(cellValue.Value);
                }
                catch (FormatException)
                {
                    Errors.Add(new ExcelDataError(ews.Name, cellValue.Address, ExcelDataErrorType.ExpectedNumeric));
                }
            }

            return ret;
        }

        private Dictionary<string, Dictionary<DateTime, decimal>> LoadIndexData()
        {
            var data = new Dictionary<string, Dictionary<DateTime, decimal>>();

            int sheetIndex = 0;
            foreach(var sheet in ep.Workbook.Worksheets)
            {
                if (sheetIndex >= dataSheetCount)
                    break;

                data[sheet.Name] = GetData(sheet);
                ++sheetIndex;
            }

            return data;
        }

        private void MatchSheetData(ExcelWorksheet ws)
        {
            var firstRow = ws.Dimension.Start.Row + 1;
            var lastRow = ws.Dimension.End.Row;

            for(int row = firstRow;row <= lastRow;++row)
            {
                var cellDay = ws.Cells[row, 1];
                var sDay = cellDay.Text;
                if (string.IsNullOrEmpty(sDay))
                    continue;

                if(!DateTime.TryParse(sDay, out DateTime day))
                {
                    Errors.Add(new ExcelDataError(ws.Name, cellDay.Address, ExcelDataErrorType.ExpectedDate));
                    continue;
                }

                if(!TryGetValueForDay(indexData[ws.Name], day, out decimal indexValue, out bool replaced))
                {
                    Errors.Add(new ExcelDataError(ws.Name, cellDay.Address, ExcelDataErrorType.MissingIndexValue));
                    continue;
                }
                int targetColumn = 5;
                ws.Cells[row, targetColumn].Value = indexValue;

                if(!TryGetValueForDay(rates, day, out decimal rate, out _))
                {
                    Errors.Add(new ExcelDataError(ws.Name, cellDay.Address, ExcelDataErrorType.MissingRate));
                    continue;
                }
                var indexValueCZK = indexValue * rate;
                int targetColumnCZ = replaced ? 4 : 3;
                ws.Cells[row, targetColumnCZ].Value = indexValueCZK;
                if (replaced)
                    SetCellColor(ref cellDay, System.Drawing.Color.Yellow);
            }
        }

        private void WriteErrors()
        {
            string errWsName = "Chyby";
            var wb = ep.Workbook;
            var errWs = wb.Worksheets.SingleOrDefault(x => x.Name == errWsName);
            if (errWs is not null)
                wb.Worksheets.Delete(errWs);

            var ws = wb.Worksheets.Add(errWsName);
            int row = 1;

            foreach (var err in Errors)
            {
                var values = new[] { err.Sheet, err.Address, err.TypeString() };
                var cellRange = ws.Cells[row, 1, row, values.Length];
                cellRange.LoadFromText(string.Join(",", values));

                if(err.Type != ExcelDataErrorType.ExpectedNumeric)
                    SetCellColor(ref cellRange, System.Drawing.Color.Red);

                ++row;
            }
        }

        private static bool TryGetValueForDay(IDictionary<DateTime, decimal> values, DateTime day, out decimal value, out bool replaced)
        {
            var date = day;
            bool hasValueForDay = values.ContainsKey(day);

            if (!hasValueForDay)
            {
                if (date < values.Keys.Min())
                {
                    (value, replaced) = (default, default);
                    return false;
                }

                while (!values.ContainsKey(date))
                {
                    date = date.AddDays(-1);
                }

                values[day] = values[date];
            }

            (value, replaced) = (values[date], !hasValueForDay);
            return true;
        }

        private static void SetCellColor(ref ExcelRange cell, System.Drawing.Color color)
        {
            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        /*
         * Returns the column address of the pos-th cell containing data in the first row of ws.
         */
        private static int GetHeaderColumn(ExcelWorksheet ws, int pos)
        {
            var cells = ws.Cells[1, 1, 1, ws.Dimension.End.Column];
            int cur = 1;
            foreach(var cell in cells)
            {
                if (cur == pos)
                    return cell.Start.Column;
                ++cur;
            }
            throw new IndexOutOfRangeException();
        }
    }

    class ExcelDataError
    {
        public ExcelDataError(string sheet, string addr, ExcelDataErrorType type)
        {
            Sheet = sheet;
            Address = addr;
            Type = type;
        }

        public string TypeString()
        {
            return Type switch
            {
                ExcelDataErrorType.ExpectedNumeric => "ExpectedNumeric",
                ExcelDataErrorType.ExpectedDate => "ExpectedDate",
                ExcelDataErrorType.MissingIndexValue => "MissingIndexValue",
                ExcelDataErrorType.MissingRate => "MissingRate",
                ExcelDataErrorType.MissingHeader => "MissingHeader",
                ExcelDataErrorType.Other => "Other",
                _ => "Other",
            };
        }

        public string Address { get; }
        public string Sheet { get; }
        public ExcelDataErrorType Type { get; }
    }

    enum ExcelDataErrorType
    {
        ExpectedNumeric,
        ExpectedDate,
        MissingIndexValue,
        MissingRate,
        MissingHeader,
        Other
    }
}
