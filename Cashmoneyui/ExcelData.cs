using System;
using System.Collections.Generic;
using System.IO;
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

        public decimal GetRate(DateTime day)
        {
            return rates[day];
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
            var missing = new HashSet<DateTime>();

            var ws = ep.Workbook.Worksheets["Kurz dolaru"];
            var firstRow = ws.Dimension.Start.Row + 1;
            var lastRow = ws.Dimension.End.Row;
            for(var row = firstRow;row < lastRow;++row)
            {
                var cellRate = ws.Cells[row, 3];
                var date = DateTime.Parse(ws.Cells[row, 2].Text);
                var rateText = cellRate.Text;
                decimal rate;
                try
                {
                    rate = decimal.Parse(rateText);
                }
                catch (FormatException)
                {
                    Errors.Add(new ExcelDataError(ws.Name, cellRate.Address, row, 3, ExcelDataErrorType.ExpectedNumeric));
                    rate = default;
                }
                if (rate != default)
                    ret[date] = rate;
                else
                    missing.Add(date);
            }

            foreach(var date in missing)
            {
                var replacementDate = date.AddDays(-1);
                while (!ret.ContainsKey(replacementDate))
                    replacementDate = replacementDate.AddDays(-1);
                ret[date] = ret[replacementDate];
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
                    break;

                var date = DateTime.Parse(cellDate.Text);
                decimal value;
                try
                {
                    value = decimal.Parse(cellValue.Text);
                }
                catch(FormatException)
                {
                    Errors.Add(new ExcelDataError(ews.Name, cellValue.Address, row, colValue, ExcelDataErrorType.ExpectedNumeric));
                    value = default;
                }
                if (value != default)
                    ret[date] = value;
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
                int targetColumn = notReplaced ? 3 : 4;
                ws.Cells[row, targetColumn].Value = indexValueCZK;
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

            while(!values.ContainsKey(date))
            {
                date = date.AddDays(-1);
            }

            return (values[date], values.ContainsKey(day));
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
