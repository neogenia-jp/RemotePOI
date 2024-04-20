using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace XlsManiSvc
{
    public class NPoiWrapper : IDisposable
    {
        private readonly ILogger<RemotePOIService> _logger;
        public DateTime CreatedAt { get; }

        public NPoiWrapper(ILogger<RemotePOIService> logger = null)
        {
            CreatedAt = DateTime.Now;
            _logger = logger;
        }

        IWorkbook book;
        ISheet sheet;

        public void LoadTemplateFromData(Stream data)
        {
            Dispose();
            book = WorkbookFactory.Create(data);
            sheet = book.GetSheetAt(0);
        }

        public void LoadTemplateFromFile(string path)
        {
            book = WorkbookFactory.Create(path);
            sheet = book.GetSheetAt(0);
        }

        public byte[] Download()
        {
            // ƒZƒ‹‚ÌŒvŽZŽ®‚ð‚·‚×‚ÄÄŒvŽZ‚³‚¹‚é
            //book.SetForceFormulaRecalculation(true);
            book.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();

            using (var mem = new MemoryStream())
            {
                book.Write(mem);
                return mem.ToArray();
            }
        }

        public void Dispose()
        {
            lock (this)
            {
                if (book != null)
                {
                    book.Close();
                    book = null;
                    sheet = null;
                }
            }
        }

        public int GetSheetCount() => book.NumberOfSheets;

        public string GetSheetName(int arg) => book.GetSheetName(arg);

        public void SelectSheetAt(int arg)
        {
            sheet = book.GetSheetAt(arg);
        }

        public void CreateSheet(string name)
        {
            sheet = book.CreateSheet(name);
        }

        public void RemoveSheetAt(int arg)
        {
            var needsMove = book.ActiveSheetIndex == arg;
            book.RemoveSheetAt(arg);
            if (needsMove)
            {
                sheet = book.GetSheetAt(0);
            }
        }

        public void SetSheetHidden(int arg, NPOI.SS.UserModel.SheetState s)
        {
            book.SetSheetHidden(arg, s);
        }

        public void SetSheetName(int arg, string name)
        {
            book.SetSheetName(arg, name);
        }

        public int CloneSheet(int arg, string name)
        {
            var s = book.CloneSheet(arg);
            var idx = book.GetSheetIndex(s.SheetName);
            SetSheetName(idx, name);
            return idx;
        }

        public void SetSheetOrder(string name, int pos)
        {
            book.SetSheetOrder(name, pos);
        }

        public CellValueTypes GetCellValueType(CellAddress addr) => sheet.GetRow(addr.Row).GetCell(addr.Col).GetCellValueType();

        public CellValue GetCellValue(CellAddress addr)
            => _GetCellValue(sheet.GetCell(addr.Row, addr.Col));
        
        private static CellValue _CreateBlankCell()
            => new CellValue()
            {
                ValueType = CellValueTypes.Unknown,
                IsBlank = true,
                StringValue = string.Empty,
            };

        private static CellValue _GetCellValue(ICell cell)
        {
            if (cell == null) return _CreateBlankCell();
            var type = cell.CellType;

            CellValue v = new CellValue();
            v.ValueType = cell.GetCellValueType();
            v.IsBlank = type == CellType.Blank;
            v.StringValue = string.Empty;

            switch (v.ValueType)
            {
                case CellValueTypes.Numeric:
                    v.NumericValue = cell.NumericCellValue;
                    v.StringValue = $"{cell.NumericCellValue}";
                    break;
                case CellValueTypes.DateTime:
                    v.DateTimeValue = cell.DateCellValue.ToTimestamp();
                    v.StringValue = cell.DateCellValue.ToLongTimeString();
                    break;
                case CellValueTypes.String:
                    v.StringValue = cell.StringCellValue;
                    break;
                case CellValueTypes.Boolean:
                    v.BoolValue = cell.BooleanCellValue;
                    v.StringValue = cell.BooleanCellValue.ToString();
                    break;
                case CellValueTypes.Error:
                    System.Console.WriteLine("row:{0} col:{1}", cell.RowIndex, cell.ColumnIndex);
                    v.ErrorValue = cell.ErrorCellValue;
                    break;
            }
            return v;
        }

        public void SetCellValue(CellAddressWithValue addrv)
        {
            var cell = sheet.GetCell(addrv.Row, addrv.Col);
            switch (addrv.Value.ValueType)
            {
                case CellValueTypes.Numeric:
                    cell.SetCellValue(addrv.Value.NumericValue);
                    break;
                case CellValueTypes.String:
                    cell.SetCellValue(addrv.Value.StringValue);
                    break;
                case CellValueTypes.Boolean:
                    cell.SetCellValue(addrv.Value.BoolValue);
                    break;
                case CellValueTypes.DateTime:
                    cell.SetCellValue(DateUtil.GetExcelDate(new DateTime(addrv.Value.DateTimeValue.Seconds)));
                    break;
                case CellValueTypes.Blank:
                    cell.SetCellType(CellType.Blank);
                    break;
                case CellValueTypes.Error:
                    cell.SetCellValue(addrv.Value.ErrorValue);
                    cell.SetCellType(CellType.Error);
                    break;
            }
        }
    }
}
