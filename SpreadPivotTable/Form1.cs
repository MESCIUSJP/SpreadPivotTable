using FarPoint.Win.Spread.Dialogs;
using GrapeCity.Spreadsheet;
using GrapeCity.Spreadsheet.PivotTables;

namespace SpreadPivotTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // IWorkbookとIWorksheetの取得
            var workbook = fpSpread1.AsWorkbook();
            var sheet1 = fpSpread1.Sheets[0].AsWorksheet();

            // ■■■ データ用シートの設定 ■■■
            // シートの設定
            sheet1.Cells.Font.Name = "メイリオ";
            sheet1.Cells.Font.Size = 11;
            sheet1.Name = "データ";
            sheet1.SetValue(0, 0, new object[,]
            {
                { "注文日", "地域", "都市", "カテゴリ", "商品", "数量" },
                { "2026-08-01", "関東", "千葉", "お菓子", "せんべい", 1120 },
                { "2026-08-01", "関東", "千葉", "お菓子", "チョコレート", 563 },
                { "2026-08-02", "関東", "東京", "お菓子", "せんべい", 1281 },
                { "2026-08-02", "関東", "東京", "お菓子", "チョコレート", 546 },
                { "2026-08-01", "関東", "千葉", "飲料", "緑茶", 326 },
                { "2026-08-02", "関東", "東京", "飲料", "緑茶", 205 },
                { "2026-08-02", "関東", "東京", "飲料", "ジュース", 186 },
                { "2026-08-01", "関西", "大阪", "お菓子", "せんべい", 1262 },
                { "2026-08-01", "関西", "大阪", "お菓子", "チョコレート", 349 },
                { "2026-08-01", "関西", "京都", "お菓子", "せんべい", 524 },
                { "2026-08-01", "関西", "京都", "お菓子", "チョコレート", 196 },
                { "2026-08-01", "関西", "大阪", "飲料", "緑茶", 363 },
                { "2026-08-01", "関西", "京都", "飲料", "緑茶", 100 },
                { "2026-08-02", "関東", "千葉", "飲料", "ジュース", 120 },
                { "2026-08-02", "関西", "大阪", "お菓子", "せんべい", 350 },
                { "2026-08-02", "関西", "京都", "飲料", "コーヒー", 180 },
                { "2026-08-02", "関東", "千葉", "お菓子", "チョコレート", 75 },
                { "2026-08-03", "関西", "大阪", "飲料", "ジュース", 210 },
                { "2026-08-03", "関東", "神奈川", "お菓子", "せんべい", 420 },
            });

            // テーブルの追加と列幅の調節
            var table = sheet1.Tables.Add(0, 0, 19, 5);
            for (var i = 0; i < sheet1.Columns.Count; i++)
            {
                sheet1.Columns[i].AutoFit();
            }

            // ■■■ ピボットテーブルの作成 ■■■
            // ピボットテーブル用シートの作成
            var sheet2 = workbook.Worksheets.Add("ピボットテーブル");
            sheet2.Activate();

            //【ステップ 1】テーブルに対応したピボットテーブルキャッシュを作成
            var pvCache = workbook.PivotCaches.Create(table);

            //【ステップ 2】ピボットテーブルを作成
            var pvTable = pvCache.CreatePivotTable(sheet2.Cells["A1"]);

            //【ステップ 3】ピボットテーブルからフィールドのコレクションを取得
            var pvFields = pvTable.PivotFields;

            //【ステップ 4】「数量」を集計用のデータフィールドとして追加
            pvTable.AddDataField(pvFields["数量"], "数量の合計", ConsolidationFunction.Sum);

            //【ステップ 5】「地域」と「都市」を行フィールドとして設定
            pvFields["地域"].Orientation = PivotFieldOrientation.Row;
            pvFields["都市"].Orientation = PivotFieldOrientation.Row;

            //【ステップ 6】「カテゴリ」と「商品」を列フィールドとして設定
            pvFields["カテゴリ"].Orientation = PivotFieldOrientation.Column;
            pvFields["商品"].Orientation = PivotFieldOrientation.Column;

            //【ステップ 7】「注文日」をフィルターフィールドとして設定
            pvFields["注文日"].Orientation = PivotFieldOrientation.Page;

            // ■■■ フィルタリングとソートの設定 ■■■
            //「カテゴリ」を"お菓子"でフィルタリング
            //pvFields["カテゴリ"].PivotFilters.Add(PivotFilterType.CaptionContains, 0, "お菓子");

            //「地域」を降順にソート
            //pvFields["地域"].AutoSortOrder = FieldSortType.Descending;

            // ■■■ ダイアログの表示 ■■■
            //this.Load += (s, ea) =>
            //{
            //    //［フィールドの設定］ダイアログの生成
            //    //var dialog = BuiltInDialogs.PivotFieldSettings(fpSpread1, pvFields["商品"]);

            //    //［値フィールドの設定］ダイアログの生成
            //    //var dialog = BuiltInDialogs.PivotValueFieldSettings(fpSpread1, pvTable.DataFields[0]);

            //    //［ピボットテーブルのフィールド］ダイアログの生成
            //    var dialog = BuiltInDialogs.PivotTableFields(fpSpread1, pvTable);

            //    //［ピボットテーブルオプション］ダイアログの生成
            //    //var dialog = BuiltInDialogs.PivotTableOptions(fpSpread1, pvTable);
            //    //if (dialog.Name == "PivotTableOptions")
            //    //{
            //    //    // サイズ変更の許可
            //    //    ((TableLayoutPanel)dialog.Controls[0]).AutoSize = false;
            //    //}

            //    // ダイアログの表示
            //    dialog.Show(this);
            //    dialog.Location = new Point(this.Right - 10, this.Top);
            //    dialog.Size = new Size(400, this.Height);
            //};

            // ■■■ レイアウトの変更 ■■■
            // コンパクト形式（既定のレイアウト）
            //pvTable.RowAxisLayout = LayoutRowType.Compact;

            // アウトライン形式
            //pvTable.RowAxisLayout = LayoutRowType.Outline;

            // 表形式
            //pvTable.RowAxisLayout = LayoutRowType.Tabular;

            // ■■■ スタイルの変更 ■■■
            // 組み込みスタイルの設定
            //var tbStyle = workbook.TableStyles[BuiltInPivotStyles.PivotStyleDark10];
            //pvTable.TableStyle = tbStyle;

            // カスタムスタイルの設定
            //var tbStyle = workbook.TableStyles.Add("CustomStyle");
            //tbStyle.ShowAsAvailablePivotTableStyle = true;
            //tbStyle[TableStyleElementType.HeaderRow].Font.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(255, 255, 255, 255);
            //tbStyle[TableStyleElementType.HeaderRow].Interior.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(255, 100, 149, 237);
            //tbStyle[TableStyleElementType.FirstSubheadingRow].Interior.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(170, 100, 149, 237);
            //tbStyle[TableStyleElementType.TotalRow].Font.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(255, 255, 255, 255);
            //tbStyle[TableStyleElementType.TotalRow].Interior.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(255, 100, 149, 237);
            //tbStyle[TableStyleElementType.WholeTable].Interior.Color =
            //    GrapeCity.Spreadsheet.Color.FromArgb(100, 100, 149, 237);
            //tbStyle[TableStyleElementType.WholeTable].Font.Name = "メイリオ";
            //tbStyle[TableStyleElementType.WholeTable].Font.Size = 11;
            //pvTable.TableStyle = tbStyle;

        }
    }
}
