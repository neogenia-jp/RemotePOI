namespace XlsManiSvc.Test;
using System;
using System.IO;
using Xunit;
using XlsManiSvc;
using Google.Protobuf.WellKnownTypes;

public class NPoiWrapperTest
{
    protected NPoiWrapper _svc;

    public NPoiWrapperTest()
    {
        _svc = new NPoiWrapper();
    }

    protected string GetSampleTemplatePath(string filename = "est_template.xls")
    {
        var pwd = Environment.CurrentDirectory;
        // Console.WriteLine(pwd); // for debug
        return Path.Combine(pwd, "../../../../../sample_template/", filename);
    }

    [Fact(DisplayName = "GetCellValueのCellタプごとのバリエーションも網羅")]
    public void GetCellValue()
    {
        _svc.LoadTemplateFromFile(GetSampleTemplatePath());

        // ブランクなセル
        var cv = _svc.GetCellValue(new CellAddress { Row = 6, Col = 0 });
        Assert.Equal(CellValueTypes.Blank, cv.ValueType);
        Assert.True(cv.IsBlank);
        Assert.Equal(0.0d, cv.NumericValue);
        Assert.Equal("", cv.StringValue);
        Assert.False(cv.BoolValue);
        Assert.Null(cv.DateTimeValue);
        Assert.Equal(0u, cv.ErrorValue);

        // 数値タイプのセル
        cv = _svc.GetCellValue(new CellAddress { Row = 10, Col = 15 });
        Assert.Equal(CellValueTypes.Numeric, cv.ValueType);
        Assert.False(cv.IsBlank);
        Assert.Equal(8000.0d, cv.NumericValue);
        Assert.Equal("8000", cv.StringValue);
        Assert.False(cv.BoolValue);
        Assert.Null(cv.DateTimeValue);
        Assert.Equal(0u, cv.ErrorValue);

        // 文字列タイプのセル
        cv = _svc.GetCellValue(new CellAddress { Row = 1, Col = 0 });
        Assert.Equal(CellValueTypes.String, cv.ValueType);
        Assert.False(cv.IsBlank);
        Assert.Equal(0.0d, cv.NumericValue);
        Assert.Equal("御　見　積　書", cv.StringValue);
        Assert.False(cv.BoolValue);
        Assert.Null(cv.DateTimeValue);
        Assert.Equal(0u, cv.ErrorValue);

        // bool タイプのセル
        // TODO

        // 日付タイプのセル
        // TODO
        
        // 式タイプのセル
        cv = _svc.GetCellValue(new CellAddress { Row = 3, Col = 6 });  // TDAY() がセットされているセル
        Assert.Equal(CellValueTypes.Formula, cv.ValueType);
        Assert.False(cv.IsBlank);
        Assert.Equal(0.0d, cv.NumericValue);
        Assert.Equal("", cv.StringValue);
        Assert.False(cv.BoolValue);
        Assert.Null(cv.DateTimeValue);
        Assert.Equal(0u, cv.ErrorValue);
    }

    [Fact(DisplayName = "SetCellValueのタイプごとの網羅")]
    public void SetCellValueのテスト()
    {
        _svc.LoadTemplateFromFile(GetSampleTemplatePath());

        // 文字列タイプのセル
        {
            _svc.SetCellValue(new CellAddressWithValue { Row = 0, Col = 1, Value = new CellValue { ValueType = CellValueTypes.String, StringValue = "あいう" } });

            var cv = _svc.GetCellValue(new CellAddress { Row = 0, Col = 1 });
            Assert.Equal(CellValueTypes.String, cv.ValueType);
            Assert.False(cv.IsBlank);
            Assert.Equal("あいう", cv.StringValue);
        }

        // 数値タイプのセル
        {
            _svc.SetCellValue(new CellAddressWithValue { Row = 7, Col = 1, Value = new CellValue { ValueType = CellValueTypes.Numeric, NumericValue = 500 } });

            var cv = _svc.GetCellValue(new CellAddress { Row = 7, Col = 1 });
            Assert.Equal(CellValueTypes.Numeric, cv.ValueType);
            Assert.False(cv.IsBlank);
            Assert.Equal("500", cv.StringValue);
            Assert.Equal(500, cv.NumericValue);
        }
        // そのあと文字列にする
        {
            _svc.SetCellValue(new CellAddressWithValue { Row = 7, Col = 1, Value = new CellValue { ValueType = CellValueTypes.String, StringValue = "600" } });

            var cv = _svc.GetCellValue(new CellAddress { Row = 7, Col = 1 });
            Assert.Equal(CellValueTypes.String, cv.ValueType);
            Assert.False(cv.IsBlank);
            Assert.Equal("600", cv.StringValue);
            Assert.Equal(0, cv.NumericValue);  // 文字列タイプなので、 NumericValue はセットされない
        }

        // 日付タイプのセル
        {
            _svc.SetCellValue(new CellAddressWithValue { Row = 3, Col = 6, Value = new CellValue { ValueType = CellValueTypes.DateTime, DateTimeValue = Google.Protobuf.WellKnownTypes.Timestamp.FromDateTime(new DateTime(2024, 12, 31, 0, 0, 0, DateTimeKind.Utc)) } });
            var cv = _svc.GetCellValue(new CellAddress { Row = 3, Col = 6 });
            Assert.Equal(CellValueTypes.DateTime, cv.ValueType);
            Assert.False(cv.IsBlank);
            Assert.Equal(new DateTime(2024, 12, 31, 0, 0, 0, DateTimeKind.Utc).ToTimestampEx(), cv.DateTimeValue);
            Assert.Equal("2024-12-31T00:00:00", cv.StringValue);
        }
        // そのあとブランクにする
        {
            _svc.SetCellValue(new CellAddressWithValue { Row = 3, Col = 6, Value = new CellValue { ValueType = CellValueTypes.Blank} });
            var cv = _svc.GetCellValue(new CellAddress { Row = 3, Col = 6 });
            Assert.Equal(CellValueTypes.Blank, cv.ValueType);
            Assert.True(cv.IsBlank);
            Assert.Equal("", cv.StringValue);
        }
        
        // boolタイプのセル
        // TODO
    }

    [Fact]
    public void Downloadの疎通()
    {
        _svc.LoadTemplateFromFile(GetSampleTemplatePath());
        var result = _svc.Download();
        Assert.NotNull(result);
    }

    [Fact(DisplayName = "シートの追加と削除")]
    public void CreateSheetとRemoveSheetのテスト()
    {
        _svc.LoadTemplateFromFile(GetSampleTemplatePath());
        Assert.Equal(1, _svc.GetSheetCount());
        Assert.Equal("お見積書", _svc.GetSheetName(0));

        // シートを１つ追加
        var sheetName = "testSheet";
        _svc.CreateSheet(sheetName);
        
        // 追加された状態を検証
        Assert.Equal(2, _svc.GetSheetCount());
        Assert.Equal("お見積書", _svc.GetSheetName(0));
        Assert.Equal(sheetName, _svc.GetSheetName(1));
        
        // もう１つ追加
        var sheetName2 = "testSheet2";
        _svc.CreateSheet(sheetName2);
        
        // 追加された状態を検証
        Assert.Equal(3, _svc.GetSheetCount());
        Assert.Equal("お見積書", _svc.GetSheetName(0));
        Assert.Equal(sheetName, _svc.GetSheetName(1));
        Assert.Equal(sheetName2, _svc.GetSheetName(2));

        // シート削除
        _svc.RemoveSheetAt(1);

        // 削除された状態を検証
        Assert.Equal(2, _svc.GetSheetCount());
        Assert.Equal("お見積書", _svc.GetSheetName(0));
        Assert.Equal(sheetName2, _svc.GetSheetName(1));
        
        // さらに削除
        _svc.RemoveSheetAt(0);

        // 削除された状態を検証
        Assert.Equal(1, _svc.GetSheetCount());
        Assert.Equal(sheetName2, _svc.GetSheetName(0));

        // さらに削除
        Assert.Throws<System.ArgumentOutOfRangeException>(()=>{
        _svc.RemoveSheetAt(0);
        });
    }

    [Fact(DisplayName = "CloneSheetのテスト", Skip = "未実装")]
    public void CloneSheetのテスト()
    {
        // FIXME

    }

    [Fact(DisplayName = "SetSheetOrderシート", Skip = "未実装")]
    public void SetSheetOrderのテスト()
    {
        // FIXME

    }

    [Fact(DisplayName = "InsertRowAt,CopyRow,ClearRowAtのテスト", Skip = "未実装")]
    public void InsertRowAtのテスト()
    {
        // FIXME

    }

    [Fact(DisplayName = "GetMaxRowNumのテスト", Skip = "未実装")]
    public void GetMaxRowNumのテスト()
    {
        // FIXME

    }

    [Fact(DisplayName = "SetCellValueで対象Rowが存在しない時のテスト")]
    public void SetCellValueで対象Rowが存在しない時のテスト()
    {
        _svc.LoadTemplateFromFile(GetSampleTemplatePath());

        Assert.Equal(399, _svc.GetMaxRowNum());

        // 最大行数を超えてセルの値を参照すると例外が出る
        Assert.Throws<System.ArgumentOutOfRangeException>(()=>{
            var cv = _svc.GetCellValue(new CellAddress { Row = 400, Col = 0 });
        });

        // 最大行数は増えていないはず
        Assert.Equal(399, _svc.GetMaxRowNum());

        // 最大行数を超えて、次の行に SetCellValue() する
        _svc.SetCellValue(new CellAddressWithValue { Row = 400, Col = 0, Value = new CellValue { ValueType = CellValueTypes.String, StringValue = "あいう" } });

        // 最大行数が増えていることを確認
        Assert.Equal(400, _svc.GetMaxRowNum());

        // セルの値が正しくセットされていることを確認
        {
            var cv = _svc.GetCellValue(new CellAddress { Row = 400, Col = 0 });
            Assert.Equal(CellValueTypes.String, cv.ValueType);
            Assert.False(cv.IsBlank);
            Assert.Equal("あいう", cv.StringValue);
        }
    }
}