using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.Util;
using Google.Protobuf.WellKnownTypes;

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
        public RecalculationPolicies recalculationPolicy { get; set; } = RecalculationPolicies.ForceEvaluate;

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
            switch (recalculationPolicy)
            {
                case RecalculationPolicies.SetFlag:
                    // Excelアプリで開いたときに再計算させるフラグを立てる
                    (book as NPOI.XSSF.UserModel.XSSFWorkbook)?.SetForceFormulaRecalculation(true);
                    break;
                case RecalculationPolicies.ForceEvaluate:
                    // セルの計算式をすべて再計算する
                    book.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
                    break;
            }

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

        public void InsertRowAt(int rownum)
        {
            sheet.CreateRow(rownum);
        }

        public void CopyRow(int sourceIndex, int targetIndex)
        {
            sheet.CopyRow(sourceIndex, targetIndex);
        }

        public void ClearRowAt(int rownum)
        {
            var row = sheet.GetRow(rownum);
            sheet.RemoveRow(row);
        }

        public void SetRowZeroHeightAt(int rownum, bool isHidden)
        {
            var row = sheet.GetRow(rownum);
            row.ZeroHeight = isHidden;
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
                    // 日付形式の書式が設定されていたら、日付型として扱う
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        v.ValueType = CellValueTypes.DateTime;
                        cell.ExtractDateTimeCellValue(v);
                        break;
                    }
                    v.NumericValue = cell.NumericCellValue;
                    v.StringValue = $"{cell.NumericCellValue}";
                    break;
                case CellValueTypes.DateTime:
                    cell.ExtractDateTimeCellValue(v);
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

        public IRow GetOrCreateRow(int rownum)
            => sheet.GetRow(rownum) ?? sheet.CreateRow(rownum);

        public ICell GetOrCreateCell(int rownum, int colnum)
        {
            var row = GetOrCreateRow(rownum);
            return row.GetCell(colnum, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        public void SetCellValue(CellAddressWithValue addrv)
        {
            var cell = GetOrCreateCell(addrv.Row, addrv.Col);
            switch (addrv.Value.ValueType)
            {
                case CellValueTypes.Numeric:
                    cell.SetCellValue(addrv.Value.NumericValue);
                    cell.SetCellType(CellType.Numeric);
                    break;
                case CellValueTypes.String:
                    cell.SetCellValue(addrv.Value.StringValue);
                    cell.SetCellType(CellType.String);
                    break;
                case CellValueTypes.Boolean:
                    cell.SetCellValue(addrv.Value.BoolValue);
                    cell.SetCellType(CellType.Boolean);
                    break;
                case CellValueTypes.DateTime:
                    cell.SetCellValue(addrv.Value.DateTimeValue.ToDateTime());
                    cell.SetCellType(CellType.Numeric);
                    // 日時データの場合は書式設定をちゃんとセットしないと、内部データの数値型のまま表示されてしまう。
                    var createHelper = book.GetCreationHelper();
                    var cellStyle = book.CreateCellStyle();
                    short style = createHelper.CreateDataFormat().GetFormat("yyyy/mm/dd h:mm");
                    cellStyle.DataFormat = style;
                    cell.CellStyle = cellStyle;
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

        public int GetMaxRowNum()
        {
            var count = 0;
            var e = sheet.GetRowEnumerator();
            while (e.MoveNext()) count++;
            return count;
        }
    }
}
