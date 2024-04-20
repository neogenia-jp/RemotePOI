using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace XlsManiSvc
{
    public static class Extensions
    {
        public static ICell GetCell(this ISheet sheet, int r, int c) => sheet.GetRow(r).GetCell(c);

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

        public static Google.Protobuf.WellKnownTypes.Timestamp ToTimestamp(this DateTime dt)
        {
            var dto = new DateTimeOffset(dt, new TimeSpan(10, 0, 0));
            return new Google.Protobuf.WellKnownTypes.Timestamp() { Seconds = dto.ToUnixTimeSeconds() };
        }
    }
}
