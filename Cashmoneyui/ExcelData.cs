using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Cashmoneyui
{
    class ExcelData
    {
        public ExcelData(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ep = new ExcelPackage(new FileInfo(filePath));
            Errors = new List<ExcelDataError>();
            rates = GetRates();
            LoadIndexData();
        }

        public void WriteMatches()
        {
            string[] sheets = { "Top stocks a Nasdaq 10Y", "Top Stock a SP 10Y", "Top Stock a DJ 10Y" };
            foreach(var sheet in sheets)
            {
                var ws = ep.Workbook.Worksheets[sheet];
                MatchSheetData(ws);
                ep.Save();
            }
        }

        ExcelPackage ep;
        SortedDictionary<DateTime, decimal> rates;
        List<ExcelDataError> Errors;
        Dictionary<string, Dictionary<DateTime, decimal>> indexData;

        private SortedDictionary<DateTime, decimal> GetRates()
        {
            var ret = new SortedDictionary<DateTime, decimal>();

            var ws = ep.Workbook.Worksheets["Kurz dolaru"];
            var firstRow = ws.Dimension.Start.Row + 1;
            var lastRow = ws.Dimension.End.Row;
            for(var row = firstRow;row < lastRow;++row)
            {
                var cellRate = ws.Cells[row, 3];
                var date = DateTime.Parse(ws.Cells[row, 2].Text);
                var rateText = cellRate.Text;
                try
                {
                    ret[date] = decimal.Parse(rateText);
                }
                catch (FormatException)
                {
                    Errors.Add(new ExcelDataError(ws.Name, cellRate.Address, row, 3, ExcelDataErrorType.ExpectedNumeric));
                }
            }

            return ret;
        }

        private Dictionary<DateTime, decimal> GetData(ExcelWorksheet ews, int colFrom)
        {
            var ret = new Dictionary<DateTime, decimal>();

            int colDate = colFrom;
            int colValue = colFrom + 1;
            var firstRow = ews.Dimension.Start.Row + 1;
            var lastRow = ews.Dimension.End.Row;

            for(var row = firstRow;row < lastRow;++row)
            {
                var cellDate = ews.Cells[row, colDate];
                var cellValue = ews.Cells[row, colValue];
                if (string.IsNullOrEmpty(cellDate.Text))
                    continue;

                var date = DateTime.Parse(cellDate.Text);
                try
                {
                    ret[date] = decimal.Parse(cellValue.Text);
                }
                catch(FormatException)
                {
                    Errors.Add(new ExcelDataError(ews.Name, cellValue.Address, row, colValue, ExcelDataErrorType.ExpectedNumeric));
                }
            }

            return ret;
        }

        private void LoadIndexData()
        {
            indexData = new Dictionary<string, Dictionary<DateTime, decimal>>();
            (string, int)[] sheets = {("Top stocks a Nasdaq 10Y", 6), ("Top Stock a SP 10Y", 7), ("Top Stock a DJ 10Y", 8)};
            foreach(var (sheet, col) in sheets)
            {
                indexData[sheet] = GetData(ep.Workbook.Worksheets[sheet], col);
            }
        }

        private void MatchSheetData(ExcelWorksheet ws)
        {
            var firstRow = ws.Dimension.Start.Row + 1;
            var lastRow = ws.Dimension.End.Row;

            for(int row = firstRow;row < lastRow;++row)
            {
                var cellDay = ws.Cells[row, 1];
                var sDay = cellDay.Text;
                if (string.IsNullOrEmpty(sDay))
                    continue;
                var day = DateTime.Parse(sDay);

                (decimal indexValue, bool notReplaced) = ValueForDay(indexData[ws.Name], day);
                (decimal rate, _) = ValueForDay(rates, day);
                var indexValueCZK = indexValue * rate;

                int targetColumn = 5;
                int targetColumnCZ = notReplaced ? 3 : 4;
                ws.Cells[row, targetColumn].Value = indexValue;
                ws.Cells[row, targetColumnCZ].Value = indexValueCZK;
                if (!notReplaced)
                {
                    cellDay.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cellDay.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                }
            }
        }

        private static (decimal, bool) ValueForDay(IDictionary<DateTime, decimal> values, DateTime day)
        {
            var date = day;
            bool hasValueForDay = values.ContainsKey(day);

            if (!hasValueForDay)
            {
                if (date < values.Keys.Min())
                    throw new IndexOutOfRangeException();

                while (!values.ContainsKey(date))
                {
                    date = date.AddDays(-1);
                }

                values[day] = values[date];
            }

            return (values[date], hasValueForDay);
        }
    }

    class ExcelDataError
    {
        public ExcelDataError(string sheet, string addr, int row, int col, ExcelDataErrorType type)
        {
            Sheet = sheet;
            Address = addr;
            Row = row;
            Column = col;
            Type = type;
        }

        string Address { get; }
        int Row { get; }
        int Column { get; }
        string Sheet { get; }
        ExcelDataErrorType Type { get; }
    }

    enum ExcelDataErrorType
    {
        ExpectedNumeric,
        MissingIndexValue,
        Other
    }
}
