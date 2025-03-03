using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace XlsManiSvc
{
    public static class Extensions
    {
        public static ICell GetCell(this ISheet sheet, int r, int c) 
        {
            var row = sheet.GetRow(r);
            if (row == null) throw new ArgumentOutOfRangeException($"Row {r} is out of range.");
            return row.GetCell(c);
        }

        public static CellValueTypes GetCellValueType(this ICell c)
           => c.CellType switch
           {
               CellType.Numeric => DateUtil.IsCellDateFormatted(c) ? CellValueTypes.DateTime : CellValueTypes.Numeric,
               CellType.String => CellValueTypes.String,
               CellType.Formula => CellValueTypes.Formula,
               CellType.Blank => CellValueTypes.Blank,
               CellType.Boolean => CellValueTypes.Boolean,
               CellType.Error => CellValueTypes.Error,
               _ => CellValueTypes.Unknown,
           };

        public static string Inspect(this CellValue cv)
        {
            switch (cv.ValueType)
            {
                case CellValueTypes.Numeric:
                    return $"Num:{cv.NumericValue}";
                case CellValueTypes.String:
                    return $"Str:{cv.StringValue}";
                case CellValueTypes.Boolean:
                    return $"Bool:{cv.BoolValue}";
                case CellValueTypes.DateTime:
                    return $"{new DateTime(cv.DateTimeValue.Seconds)} (epoc: {cv.DateTimeValue.Seconds})";
                case CellValueTypes.Formula:
                    return $"Fml:{cv.StringValue}";
                case CellValueTypes.Blank:
                    return "(Blank)";
                case CellValueTypes.Error:
                    return $"Error:{cv.ErrorValue}";
            }
            return "(Unknown)";
        }

        /// <summary>
        /// cell から日時データを取り出して、v にセットする
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="v"></param>
        public static DateTime? ExtractDateTimeCellValue(this ICell cell, CellValue v = null)
        {
            if (!cell.DateCellValue.HasValue) return null;
            
        var dt =  new DateTime(cell.DateCellValue.Value.Ticks, DateTimeKind.Utc);  // 強制的にUTCとして扱う
                if (v != null)
                {
                v.DateTimeValue = dt.ToTimestampEx();
                v.StringValue = dt.ToString("s");  // ISO8601形式
                }
                return dt;
        }

        public static Google.Protobuf.WellKnownTypes.Timestamp ToTimestampEx(this DateTime dt)
        {
            if (dt.Kind == DateTimeKind.Utc)
            {
                // UTC だった場合はそのまま変換
                return new Google.Protobuf.WellKnownTypes.Timestamp() { Seconds = dt.Ticks };
            }
            var dto = new DateTime(dt.Ticks, DateTimeKind.Utc);
            return new Google.Protobuf.WellKnownTypes.Timestamp() { Seconds = dto.Ticks };
        }
    }
}
