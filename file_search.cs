// ==============================================================================
// file_search.cs
// 預貯金照会結果 検索ツール（WPF GUI版）
//
// 【使用方法】
// build.bat を実行して file_search.exe を生成し、ダブルクリックで起動する。
// 宛名番号を入力して、設定で指定したフォルダ配下から .xlsm ファイルを検索する。
// 検索結果から「ファイルを開く」または「差押リストに追加」を選択できる。
// ★お気に入りフォルダ検索モードでは、ユーザー設定のフォルダを対象に
// 全件表示やファイル名検索が可能。
// 「差押リストに追加」を選ぶと、ファイル内容から預金/生保を自動判定し、
// deposit_seizure_list.exe または insurance_seizure_list.exe を呼び出す。
//
// 【ビルド方法】
// build.bat を実行（.NET Framework 4.0 の csc.exe を使用）
// ※ お気に入りフォルダの参照ダイアログに System.Windows.Forms.dll が必要
//
// 【必要ファイル（同じフォルダに配置）】
// ＜必須＞
// - file_search.cs  （ソースコード）
// - file_search_config.json （設定ファイル）
// - file_search.ico （アプリケーションアイコン）
// - build.bat （ビルドスクリプト）
// ＜任意＞
// - update_notice.json （アップデート告知。配置時のみ初回起動時に表示）
// ==============================================================================

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Xml;

// ==============================================================
// 設定
// ==============================================================

public class AppConfig
{
    public string[] SearchFolders { get; set; }
    public string DepositScript { get; set; }    // null の場合は預金差押リスト追加が無効
    public string InsuranceScript { get; set; }  // null の場合は生保差押リスト追加が無効
    public string SeizureManagerScript { get; set; }  // null の場合は管理ツール起動ボタンが非表示
    public string DepositDetectCell { get; set; }    // 預金/生保 自動判定セル（例: "A2"）
    public string DepositDetectValue { get; set; }   // 自動判定値（例: "回答書（単票）"）
    public int SearchHistoryMax { get; set; }    // 検索履歴の上限件数
    public int FileHistoryMax { get; set; }      // ファイル履歴の上限件数
    public int SeizureLogMax { get; set; }       // 差押登録ログの上限件数

    // 差押リスト追加機能の有効判定（預金または生保の少なくとも一方が設定済み）
    public bool HasSeizureScript { get { return DepositScript != null || InsuranceScript != null; } }

    public AppConfig()
    {
        SearchFolders = new string[0];
        DepositScript = null;
        InsuranceScript = null;
        SeizureManagerScript = null;
        DepositDetectCell = null;
        DepositDetectValue = null;
        SearchHistoryMax = 5;
        FileHistoryMax = 10;
        SeizureLogMax = 15;
    }

    // file_search_config.json を読み込み、AppConfig を返す
    // 外部 JSON ライブラリが使えない（LGWAN 環境）ため手動パース
    public static AppConfig Load(string configPath)
    {
        var config = new AppConfig();
        if (!File.Exists(configPath)) return config;

        try
        {
            var json = File.ReadAllText(configPath, Encoding.UTF8);
            var lines = json.Split('\n');
            var folders = new List<string>();
            bool inArray = false;

            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim().TrimEnd(',').Trim();

                if (line.Contains("\"searchFolder\""))
                {
                    if (line.Contains("["))
                    {
                        inArray = true;
                        if (line.Contains("]"))
                        {
                            // 1行に配列がまるごと収まるケース
                            inArray = false;
                            var start = line.IndexOf('[');
                            var end = line.IndexOf(']');
                            var inner = line.Substring(start + 1, end - start - 1);
                            foreach (var part in inner.Split(','))
                            {
                                var p = part.Trim().Trim('"');
                                if (!string.IsNullOrWhiteSpace(p))
                                    folders.Add(p.Replace("\\\\", "\\"));
                            }
                        }
                    }
                    else
                    {
                        // 文字列（単一フォルダ）
                        var colonIdx = line.IndexOf(':');
                        if (colonIdx >= 0)
                        {
                            var val = line.Substring(colonIdx + 1).Trim().Trim('"');
                            if (!string.IsNullOrWhiteSpace(val))
                                folders.Add(val.Replace("\\\\", "\\"));
                        }
                    }
                }
                else if (inArray)
                {
                    // 配列の中身を1行ずつ読み取る
                    if (line.Contains("]"))
                    {
                        inArray = false;
                        var val = line.Replace("]", "").Trim().Trim('"');
                        if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("//") && !val.StartsWith("_"))
                            folders.Add(val.Replace("\\\\", "\\"));
                    }
                    else
                    {
                        var val = line.Trim().Trim('"');
                        if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("//") && !val.StartsWith("_"))
                            folders.Add(val.Replace("\\\\", "\\"));
                    }
                }
                else if (line.Contains("\"depositSeizureListScript\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var val = line.Substring(colonIdx + 1).Trim().Trim('"').Replace("\\\\", "\\");
                        if (!string.IsNullOrWhiteSpace(val) && File.Exists(val))
                            config.DepositScript = val;
                    }
                }
                else if (line.Contains("\"seizureListManagerScript\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var val = line.Substring(colonIdx + 1).Trim().Trim('"').Replace("\\\\", "\\");
                        if (!string.IsNullOrWhiteSpace(val) && File.Exists(val))
                            config.SeizureManagerScript = val;
                    }
                }
                else if (line.Contains("\"insuranceSeizureListScript\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var val = line.Substring(colonIdx + 1).Trim().Trim('"').Replace("\\\\", "\\");
                        if (!string.IsNullOrWhiteSpace(val) && File.Exists(val))
                            config.InsuranceScript = val;
                    }
                }
                else if (line.Contains("\"depositDetectCell\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var val = line.Substring(colonIdx + 1).Trim().Trim('"').TrimEnd(',').Trim();
                        if (!string.IsNullOrWhiteSpace(val))
                            config.DepositDetectCell = val;
                    }
                }
                else if (line.Contains("\"depositDetectValue\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        var val = line.Substring(colonIdx + 1).Trim().Trim('"').TrimEnd(',').Trim();
                        if (!string.IsNullOrWhiteSpace(val))
                            config.DepositDetectValue = val;
                    }
                }
                else if (line.Contains("\"searchHistoryMax\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        int val;
                        if (int.TryParse(line.Substring(colonIdx + 1).Trim().TrimEnd(','), out val) && val > 0)
                            config.SearchHistoryMax = val;
                    }
                }
                else if (line.Contains("\"fileHistoryMax\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        int val;
                        if (int.TryParse(line.Substring(colonIdx + 1).Trim().TrimEnd(','), out val) && val > 0)
                            config.FileHistoryMax = val;
                    }
                }
                else if (line.Contains("\"seizureLogMax\""))
                {
                    var colonIdx = line.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        int val;
                        if (int.TryParse(line.Substring(colonIdx + 1).Trim().TrimEnd(','), out val) && val > 0)
                            config.SeizureLogMax = val;
                    }
                }
            }

            // 存在するフォルダのみを採用
            config.SearchFolders = folders.Where(f => Directory.Exists(f)).ToArray();
        }
        catch { /* パース失敗時はデフォルト値で返す */ }

        return config;
    }
}

// ==============================================================
// データモデル
// ==============================================================

// 検索結果1件を表す（ListView の各行に対応）
public class SearchResultItem
{
    public string FileName { get; set; }       // ファイル名（拡張子付き）
    public string FolderPath { get; set; }     // searchFolder からの相対ディレクトリパス
    public DateTime LastWriteTime { get; set; } // ファイルの最終更新日時（ソート用）
    public string LastWrite { get; set; }      // 表示用の更新日時文字列 (MM/dd HH:mm)
    public long SizeBytes { get; set; }        // ファイルサイズ（ソート用）
    public string Size { get; set; }           // 表示用のサイズ文字列 (KB/MB)
    public string FullPath { get; set; }       // ファイルのフルパス
}

// ファイル履歴1件を表す（サイドパネル + JSON 永続化）
public class FileHistoryEntry
{
    public string Path { get; set; }           // ファイルのフルパス
    public DateTime OpenedAt { get; set; }     // ユーザーがファイルを開いた日時
}

// アップデート告知（update_notice.json のデシリアライズ用）
public class UpdateNotice
{
    public string Version { get; set; }
    public List<NoticeFeature> Features { get; set; }
    public UpdateNotice() { Features = new List<NoticeFeature>(); }
}
public class NoticeFeature
{
    public string Title { get; set; }
    public string Description { get; set; }
    public List<NoticeSection> Sections { get; set; }
    public NoticeFeature() { Sections = new List<NoticeSection>(); }
}
public class NoticeSection
{
    public string Heading { get; set; }
    public List<string> Items { get; set; }
    public NoticeSection() { Items = new List<string>(); }
}

// 差押登録ログ1件を表す（サイドパネルの差押登録タブ用）
public class SeizureLogEntry
{
    public string AddressNumber { get; set; }   // 宛名番号
    public string Name { get; set; }            // 氏名
    public string InstitutionName { get; set; } // 金融機関名
    public string ExecutionDate { get; set; }   // 執行日（yyyy-MM-dd）
    public string DocumentNumber { get; set; }  // 文書番号
    public DateTime RegisteredAt { get; set; }  // 登録日時
}

// ==============================================================
// メインアプリケーション
// ==============================================================

public class FileSearchApp : Application
{
    // --- 設定 ---
    private AppConfig config;
    private string exeDir;                     // exe と同階層のフォルダパス

    // --- UI要素 ---
    private Window window;
    private TextBox searchBox;
    private TextBlock searchPlaceholder;
    private Button searchButton;
    private Button clearButton;                // 検索ボックスの × クリアボタン
    private ListView resultList;
    private Button openButton;
    private Button addButton;                  // 差押ツール未設定時は Collapsed
    private Button managerButton;              // seizureManagerScript 未設定時は Collapsed
    private ListBox fileHistoryList;
    private ListBox searchHistoryList;          // サイドパネル検索履歴タブ
    private Border tabHistory, tabFile, tabSeizure;  // サイドパネルのタブヘッダー
    private TextBlock tabHistoryText, tabFileText, tabSeizureText;
    private ListBox seizureLogList;                   // サイドパネル差押登録タブ
    private List<SeizureLogEntry> seizureLog = new List<SeizureLogEntry>();
    private TextBlock statusLeft;              // フッター左（件数表示）
    private TextBlock statusRight;             // フッター右（検索フォルダ表示）

    // --- お気に入りフォルダUI要素 ---
    private Button starButton;                           // ★お気に入りフォルダ検索トグル
    private Grid favOverlay;                             // お気に入り設定オーバーレイ
    private StackPanel favSlotPanel;                     // フォルダスロット一覧
    private TextBox favPathInput;                        // フォルダパス入力欄
    private TextBlock favSectionBadge;                   // セクションヘッダーの★バッジ

    // --- 状態 ---
    private List<string> searchHistory = new List<string>();
    private List<FileHistoryEntry> fileHistory = new List<FileHistoryEntry>();
    private List<SearchResultItem> currentResults = new List<SearchResultItem>();
    private string currentSortColumn = "LastWriteTime";  // 現在のソート列
    private bool currentSortAscending = false;           // true=昇順, false=降順
    private bool persistHistory = true;                  // 永続化が有効か（読み書き失敗時に false に切り替わる）
    private string activeTab = "history";                 // サイドパネルのアクティブタブ（"history" / "file" / "seizure"）
    private FrameworkElement resultItemsPanel;            // 検索結果のデータ行エリア（フェードアニメーション用）
    private BackgroundWorker searchWorker;                // 検索処理の非同期実行用
    private Border searchOverlay;                        // 検索中のローディングオーバーレイ
    private RotateTransform spinnerRotation;              // スピナーの回転トランスフォーム

    // --- お気に入りフォルダ検索 ---
    private bool isFavoriteMode = false;                  // お気に入りフォルダ検索モードか
    private List<string> favoriteFolders = new List<string>();
    private const int FAVORITE_FOLDERS_MAX = 5;           // お気に入りフォルダの上限
    private bool pendingFavoriteRefresh = false;          // キャンセル後の全件再走査を保留中か
    private string[] favFoldersSnapshot;                  // オーバーレイ表示時のスナップショット（変更検知用）

    // --- アップデート告知 ---
    private string lastSeenVersion = "";                  // ユーザーが最後に確認した告知バージョン
    private string pendingNoticeVersion = "";             // 表示中の告知のバージョン
    private Grid noticeOverlay;                           // アップデート告知オーバーレイ
    private StackPanel noticeContent;                     // 告知内容の動的描画エリア

    // --- キャッシュ済みブラシ（MakeIcon・UpdateFileHistoryUI 用） ---
    private static readonly SolidColorBrush BrushIconGray    = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#555555"));
    private static readonly SolidColorBrush BrushIconGreen   = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#107C41"));
    private static readonly SolidColorBrush BrushFolderFill  = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8EEF4"));
    private static readonly SolidColorBrush BrushAccent      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#005FB8"));
    private static readonly SolidColorBrush BrushSecondary   = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#999999"));
    private static readonly SolidColorBrush BrushTabInactive = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#777777"));
    private static readonly SolidColorBrush BrushFooter      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#666666"));
    private static readonly SolidColorBrush BrushError       = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D32F2F"));
    private static readonly SolidColorBrush BrushStarOn      = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E8F0FE"));
    private static readonly SolidColorBrush BrushBorderLight = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#D0D0D0"));
    private static readonly SolidColorBrush BrushBorderNormal = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#E0E0E0"));
    private static readonly SolidColorBrush BrushSlotEmpty   = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FAFAFA"));
    private static readonly SolidColorBrush BrushHoverBg     = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#F0F0F0"));

    // --- ソート列名のマッピング（UpdateSortIndicators 用） ---
    private static readonly Dictionary<string, string> SortColumnNames = new Dictionary<string, string>
    {
        { "FileName", "ファイル名" },
        { "FolderPath", "フォルダ" },
        { "LastWriteTime", "更新日" },
        { "Size", "サイズ" }
    };

    // ==============================================================
    // エントリポイント
    // ==============================================================

    [STAThread]
    public static void Main()
    {
        // ブラシを Freeze して描画パフォーマンスを向上
        BrushIconGray.Freeze();
        BrushIconGreen.Freeze();
        BrushFolderFill.Freeze();
        BrushAccent.Freeze();
        BrushSecondary.Freeze();
        BrushTabInactive.Freeze();
        BrushFooter.Freeze();
        BrushError.Freeze();
        BrushStarOn.Freeze();
        BrushBorderLight.Freeze();
        BrushBorderNormal.Freeze();
        BrushSlotEmpty.Freeze();
        BrushHoverBg.Freeze();

        var app = new FileSearchApp();
        app.Run();
    }

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var configPath = System.IO.Path.Combine(exeDir, "file_search_config.json");

        // --- 設定ファイルの存在チェック ---
        if (!File.Exists(configPath))
        {
            MessageBox.Show(
                "設定ファイルが見つかりません。\n\n" + configPath,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        config = AppConfig.Load(configPath);

        // --- searchFolder のチェック（必須・存在確認） ---
        if (config.SearchFolders.Length == 0)
        {
            MessageBox.Show(
                "searchFolder が設定されていないか、指定されたフォルダが存在しません。\nfile_search_config.json を確認してください。",
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(1);
            return;
        }

        // 差押ツールのパスチェックは AppConfig.Load 内で実施済み
        // 存在しない場合は null（各機能が個別に無効化される）

        // --- 履歴の読み込み ---
        LoadHistory();

        // --- ウィンドウの構築 ---
        window = BuildWindow();
        FindControls();
        SetupEvents();
        InitializeUI();

        // 起動後に検索ボックスにフォーカス + アップデート告知チェック
        window.ContentRendered += delegate
        {
            searchBox.Focus();
            CheckUpdateNotice();
        };
        window.Show();
    }

    // ==============================================================
    // UI要素の取得
    // ==============================================================

    // BuildWindow で生成した XAML ツリーから x:Name で各要素を取得
    private void FindControls()
    {
        searchBox         = (TextBox)window.FindName("SearchBox");
        searchPlaceholder = (TextBlock)window.FindName("SearchPlaceholder");
        searchButton      = (Button)window.FindName("SearchButton");
        clearButton       = (Button)window.FindName("ClearButton");
        resultList        = (ListView)window.FindName("ResultList");
        openButton        = (Button)window.FindName("OpenButton");
        addButton         = (Button)window.FindName("AddButton");
        managerButton     = (Button)window.FindName("ManagerButton");
        fileHistoryList   = (ListBox)window.FindName("FileHistoryList");
        searchHistoryList = (ListBox)window.FindName("SearchHistoryList");
        tabHistory        = (Border)window.FindName("TabHistory");
        tabFile           = (Border)window.FindName("TabFile");
        tabSeizure        = (Border)window.FindName("TabSeizure");
        tabHistoryText    = (TextBlock)window.FindName("TabHistoryText");
        tabFileText       = (TextBlock)window.FindName("TabFileText");
        tabSeizureText    = (TextBlock)window.FindName("TabSeizureText");
        seizureLogList    = (ListBox)window.FindName("SeizureLogList");
        statusLeft        = (TextBlock)window.FindName("StatusLeft");
        statusRight       = (TextBlock)window.FindName("StatusRight");
        searchOverlay     = (Border)window.FindName("SearchOverlay");

        // スピナーの RotateTransform を取得（回転アニメーションの開始/停止を制御するため）
        var spinnerPath = (FrameworkElement)window.FindName("SpinnerPath");
        if (spinnerPath != null)
            spinnerRotation = spinnerPath.RenderTransform as RotateTransform;

        // お気に入りフォルダ関連
        starButton       = (Button)window.FindName("StarButton");
        favOverlay        = (Grid)window.FindName("FavOverlay");
        favSlotPanel      = (StackPanel)window.FindName("FavSlotPanel");
        favPathInput      = (TextBox)window.FindName("FavPathInput");
        favSectionBadge   = (TextBlock)window.FindName("FavSectionBadge");

        // アップデート告知関連
        noticeOverlay    = (Grid)window.FindName("NoticeOverlay");
        noticeContent    = (StackPanel)window.FindName("NoticeContent");
    }

    // ==============================================================
    // 初期化
    // ==============================================================

    // UI の初期状態を設定
    private void InitializeUI()
    {
        // 差押ツール未設定時は「差押リストに追加」ボタンは XAML 側で非表示済み

        // フッター右: 検索フォルダ表示
        if (config.SearchFolders.Length == 1)
        {
            statusRight.Text = config.SearchFolders[0];
            statusRight.ToolTip = config.SearchFolders[0];
        }
        else
        {
            // 複数フォルダ時はツールチップにフルパス一覧を表示
            statusRight.Text = config.SearchFolders.Length + " 個のフォルダを検索中";
            statusRight.ToolTip = string.Join("\n", config.SearchFolders);
        }

        // 履歴の表示
        UpdateFileHistoryUI();
        UpdateSearchHistoryUI();

        // ボタン初期状態（選択なし → 無効）
        openButton.IsEnabled = false;
        addButton.IsEnabled = false;

        // お気に入りフォルダスロットの初期描画
        UpdateFavoriteUI();

        // 差押登録ログの初期読み込み
        LoadSeizureLog();
        UpdateSeizureLogUI();
    }

    // ==============================================================
    // イベントハンドラ
    // ==============================================================

    private void SetupEvents()
    {
        // --- 検索ボックス ---

        // プレースホルダ・クリアボタンの表示切替
        searchBox.TextChanged += delegate
        {
            bool hasText = searchBox.Text.Length > 0;
            searchPlaceholder.Visibility = hasText ? Visibility.Collapsed : Visibility.Visible;
            clearButton.Visibility = hasText ? Visibility.Visible : Visibility.Collapsed;
        };

        // Enter で検索実行
        searchBox.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                ExecuteSearch(searchBox.Text);
                e.Handled = true;
            }
        };

        // --- 検索ボタン ---
        searchButton.Click += delegate { ExecuteSearch(searchBox.Text); };

        // --- クリアボタン（検索ボックス内 ×）---
        clearButton.Click += delegate
        {
            if (searchWorker.IsBusy)
                searchWorker.CancelAsync();
            searchBox.Text = "";
            searchOverlay.BeginAnimation(UIElement.OpacityProperty, null);
            searchOverlay.Visibility = Visibility.Collapsed;
            statusLeft.Foreground = BrushFooter;
            searchButton.IsEnabled = true;
            openButton.IsEnabled = false;
            addButton.IsEnabled = false;

            if (isFavoriteMode)
            {
                // お気に入りモード中は全件再走査
                // 検索中だった場合は Completed で再走査する
                if (searchWorker.IsBusy)
                    pendingFavoriteRefresh = true;
                else
                    ExecuteSearch("");
            }
            else
            {
                currentResults.Clear();
                resultList.ItemsSource = null;
                statusLeft.Text = "";
            }

            searchBox.Focus();
        };

        // --- ★お気に入りフォルダボタン ---
        starButton.Click += delegate { ToggleFavoriteMode(); };
        starButton.MouseRightButtonUp += delegate(object s, MouseButtonEventArgs e)
        {
            ShowFavoriteOverlay();
            e.Handled = true;
        };

        // --- お気に入り設定オーバーレイ ---
        var favCloseButton = (Button)window.FindName("FavCloseButton");
        var favBrowseButton = (Button)window.FindName("FavBrowseButton");
        var favAddButton = (Button)window.FindName("FavAddButton");

        favCloseButton.Click += delegate { CloseFavoriteOverlay(); };
        favBrowseButton.Click += delegate
        {
            // FolderBrowserDialog でフォルダを選択しパス入力欄に反映
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.Description = "お気に入りに追加するフォルダを選択してください";
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                favPathInput.Text = dlg.SelectedPath;
        };
        favAddButton.Click += delegate { AddFavoriteFolder(); };
        favPathInput.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.Key == Key.Return) { AddFavoriteFolder(); e.Handled = true; }
        };

        // オーバーレイの暗転背景クリックで閉じる（カード内クリックは無視）
        var favCard = (Border)window.FindName("FavCard");
        favOverlay.MouseDown += delegate(object s, MouseButtonEventArgs e)
        {
            var hit = e.OriginalSource as DependencyObject;
            while (hit != null)
            {
                if (hit == favCard) return;  // カード内のクリックは無視
                hit = VisualTreeHelper.GetParent(hit);
            }
            CloseFavoriteOverlay();
        };

        // --- アップデート告知オーバーレイ ---
        var noticeCloseButton = (Button)window.FindName("NoticeCloseButton");
        noticeCloseButton.Click += delegate { CloseNoticeOverlay(); };

        // --- サイドパネルのタブ切替 ---
        tabHistory.MouseDown += delegate { SwitchTab("history"); };
        tabFile.MouseDown += delegate { SwitchTab("file"); };
        tabSeizure.MouseDown += delegate { SwitchTab("seizure"); };

        // --- 検索履歴タブ: 左クリックで即検索実行 ---
        searchHistoryList.PreviewMouseUp += delegate(object s, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            if (searchWorker.IsBusy) return;
            var idx = searchHistoryList.SelectedIndex;
            if (idx < 0 || idx >= searchHistory.Count) return;
            var query = searchHistory[idx];
            searchBox.Text = query;
            searchBox.CaretIndex = query.Length;
            ExecuteSearch(query);
        };

        // --- 検索結果の選択（排他選択: サイドパネルの選択を解除） ---
        resultList.SelectionChanged += delegate
        {
            if (resultList.SelectedItem != null)
            {
                fileHistoryList.SelectedIndex = -1;
                searchHistoryList.SelectedIndex = -1;
            }
            UpdateButtonState();
        };

        // ダブルクリックでファイルを開く
        resultList.MouseDoubleClick += delegate
        {
            var item = resultList.SelectedItem as SearchResultItem;
            if (item != null) OpenFile(item.FullPath);
        };

        // --- ファイル履歴タブの選択（排他選択: メインパネルの選択を解除） ---
        fileHistoryList.SelectionChanged += delegate
        {
            if (fileHistoryList.SelectedItem != null)
            {
                resultList.SelectedIndex = -1;
                searchHistoryList.SelectedIndex = -1;
            }
            UpdateButtonState();
        };

        // ダブルクリックでファイルを開く
        fileHistoryList.MouseDoubleClick += delegate
        {
            var entry = GetSelectedFileHistoryEntry();
            if (entry != null) OpenFile(entry.Path);
        };

        // --- 検索履歴タブの選択（排他選択、ボタンには影響しない） ---
        searchHistoryList.SelectionChanged += delegate
        {
            if (searchHistoryList.SelectedItem != null)
            {
                resultList.SelectedIndex = -1;
                fileHistoryList.SelectedIndex = -1;
            }
        };

        // --- アクションボタン ---
        openButton.Click += delegate { OpenSelectedFiles(); };
        addButton.Click += delegate { AddSelectedToSeizureList(); };
        managerButton.Click += delegate { LaunchSeizureListManager(); };

        // --- 列ヘッダーソート ---
        resultList.AddHandler(GridViewColumnHeader.ClickEvent,
            new RoutedEventHandler(OnColumnHeaderClick));

        // --- コンテキストメニュー ---
        SetupContextMenu();

        // --- キーボードショートカット ---
        window.InputBindings.Add(new KeyBinding(
            new RelayCommand(p => OpenSelectedFiles()),
            new KeyGesture(Key.O, ModifierKeys.Control)));

        window.InputBindings.Add(new KeyBinding(
            new RelayCommand(p => {
                var query = searchBox.Text;
                if (!string.IsNullOrWhiteSpace(query) || isFavoriteMode)
                    ExecuteSearch(query ?? "");
            }),
            new KeyGesture(Key.F5)));

        window.InputBindings.Add(new KeyBinding(
            new RelayCommand(p => DeleteSelectedHistory()),
            new KeyGesture(Key.Delete)));

        window.InputBindings.Add(new KeyBinding(
            new RelayCommand(p => { if (favOverlay.Visibility == Visibility.Visible) CloseFavoriteOverlay(); }),
            new KeyGesture(Key.Escape)));

        // --- 空白領域クリックで選択・フォーカスを解除 ---
        window.MouseDown += delegate(object s, MouseButtonEventArgs e)
        {
            // 操作対象の要素上のクリックは無視
            var hit = e.OriginalSource as DependencyObject;
            while (hit != null)
            {
                if (hit is ListViewItem || hit is ListBoxItem
                    || hit is TextBox || hit is Button) return;
                hit = VisualTreeHelper.GetParent(hit);
            }
            Keyboard.ClearFocus();
            resultList.SelectedIndex = -1;
            fileHistoryList.SelectedIndex = -1;
            searchHistoryList.SelectedIndex = -1;
            UpdateButtonState();
        };

        // --- 列幅自動調整（ウィンドウリサイズ時にファイル名列が残幅を埋める） ---
        resultList.SizeChanged += delegate { AdjustColumnWidths(); };
        resultList.Loaded += delegate
        {
            AdjustColumnWidths();
            // オーバーレイを列ヘッダーの下に配置（ヘッダーの実高さを取得して Margin に反映）
            var headerRow = FindVisualChild<GridViewHeaderRowPresenter>(resultList);
            if (headerRow != null)
                searchOverlay.Margin = new Thickness(0, headerRow.ActualHeight, 0, 0);
        };

        // --- スピナーアニメーション: オーバーレイの表示/非表示に連動 ---
        searchOverlay.IsVisibleChanged += delegate(object s, DependencyPropertyChangedEventArgs dpce)
        {
            if ((bool)dpce.NewValue) StartSpinner();
            else StopSpinner();
        };

        // --- 検索処理の非同期ワーカー ---
        searchWorker = new BackgroundWorker { WorkerSupportsCancellation = true };
        searchWorker.DoWork += SearchWorker_DoWork;
        searchWorker.RunWorkerCompleted += SearchWorker_Completed;
    }

    // ファイル名列が残りスペースを埋めるよう列幅を再計算
    private void AdjustColumnWidths()
    {
        var gv = resultList.View as GridView;
        if (gv == null || gv.Columns.Count < 4) return;

        double totalWidth = resultList.ActualWidth
            - SystemParameters.VerticalScrollBarWidth - 10;
        if (totalWidth <= 0) return;

        // フォルダ、更新日、サイズは固定幅、ファイル名が可変
        double folderW = 180;
        double dateW = 130;
        double sizeW = 80;
        double fileNameW = totalWidth - folderW - dateW - sizeW;
        if (fileNameW < 150) fileNameW = 150;

        gv.Columns[0].Width = fileNameW;
        gv.Columns[1].Width = folderW;
        gv.Columns[2].Width = dateW;
        gv.Columns[3].Width = sizeW;
    }

    // Dispatcher の Render 優先度で空処理を実行し、保留中の描画を強制反映させる
    private void ForceRender()
    {
        window.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Render, new Action(() => { }));
    }

    // ビジュアルツリーから指定型の子要素を深さ優先で検索
    private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            if (child is T) return (T)child;
            var found = FindVisualChild<T>(child);
            if (found != null) return found;
        }
        return null;
    }

    // スピナー回転アニメーションを開始（RepeatBehavior.Forever で無限回転）
    private void StartSpinner()
    {
        if (spinnerRotation == null) return;
        var anim = new DoubleAnimation();
        anim.By = 360;
        anim.Duration = new Duration(TimeSpan.FromSeconds(1));
        anim.RepeatBehavior = RepeatBehavior.Forever;
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, anim);
    }

    // スピナー回転アニメーションを停止
    private void StopSpinner()
    {
        if (spinnerRotation == null) return;
        spinnerRotation.BeginAnimation(RotateTransform.AngleProperty, null);
    }

    // ==============================================================
    // コンテキストメニュー
    // ==============================================================

    // 検索結果パネルとサイドパネルにそれぞれ右クリックメニューを設定
    private void SetupContextMenu()
    {
        // --- 検索結果用 ---
        var resultMenu = new ContextMenu();
        var resultOpen = new MenuItem { Header = "ファイルを開く", Icon = MakeIcon(IconKind.Open) };
        resultOpen.Click += delegate { OpenSelectedFiles(); };
        resultMenu.Items.Add(resultOpen);

        if (config.HasSeizureScript)
        {
            var resultAdd = new MenuItem { Header = "差押リストに追加", Icon = MakeIcon(IconKind.Plus) };
            resultAdd.Click += delegate { AddSelectedToSeizureList(); };
            resultMenu.Items.Add(resultAdd);
        }

        resultMenu.Items.Add(new Separator());
        var resultFolder = new MenuItem { Header = "フォルダを開く", Icon = MakeIcon(IconKind.Folder) };
        resultFolder.Click += delegate { OpenSelectedFolder(); };
        resultMenu.Items.Add(resultFolder);

        resultList.ContextMenu = resultMenu;

        // --- サイドパネル用 ---
        var histMenu = new ContextMenu();
        var histOpen = new MenuItem { Header = "ファイルを開く", Icon = MakeIcon(IconKind.Open) };
        histOpen.Click += delegate
        {
            var entry = GetSelectedFileHistoryEntry();
            if (entry != null) OpenFile(entry.Path);
        };
        histMenu.Items.Add(histOpen);

        if (config.HasSeizureScript)
        {
            var histAdd = new MenuItem { Header = "差押リストに追加", Icon = MakeIcon(IconKind.Plus) };
            histAdd.Click += delegate
            {
                var entry = GetSelectedFileHistoryEntry();
                if (entry != null) AddToSeizureList(entry.Path);
            };
            histMenu.Items.Add(histAdd);
        }

        histMenu.Items.Add(new Separator());
        var histFolder = new MenuItem { Header = "フォルダを開く", Icon = MakeIcon(IconKind.Folder) };
        histFolder.Click += delegate
        {
            var entry = GetSelectedFileHistoryEntry();
            if (entry != null) OpenFolder(entry.Path);
        };
        histMenu.Items.Add(histFolder);

        histMenu.Items.Add(new Separator());
        var histDelete = new MenuItem { Header = "削除" };
        histDelete.Click += delegate { DeleteSelectedHistory(); };
        histMenu.Items.Add(histDelete);
        var histClearAll = new MenuItem { Header = "全て削除" };
        histClearAll.Click += delegate { ClearAllHistory("file"); };
        histMenu.Items.Add(histClearAll);

        fileHistoryList.ContextMenu = histMenu;

        // --- 検索履歴タブ用 ---
        var searchHistMenu = new ContextMenu();
        var searchHistDelete = new MenuItem { Header = "削除" };
        searchHistDelete.Click += delegate { DeleteSelectedHistory(); };
        searchHistMenu.Items.Add(searchHistDelete);
        var searchHistClearAll = new MenuItem { Header = "全て削除" };
        searchHistClearAll.Click += delegate { ClearAllHistory("history"); };
        searchHistMenu.Items.Add(searchHistClearAll);

        searchHistoryList.ContextMenu = searchHistMenu;
    }

    // ==============================================================
    // アイコン生成（コンテキストメニュー用）
    // ==============================================================

    private enum IconKind { Search, Open, Plus, Folder }

    // XAML Path で描画するアイコンを生成（外部フォント・画像への依存なし）
    // ブラシはキャッシュ済みの静的フィールドを使用
    private UIElement MakeIcon(IconKind kind)
    {
        var canvas = new Canvas { Width = 16, Height = 16 };

        switch (kind)
        {
            case IconKind.Search:
                var searchEllipse = new System.Windows.Shapes.Ellipse
                {
                    Width = 9, Height = 9, Stroke = BrushIconGray,
                    StrokeThickness = 1.5, Fill = Brushes.Transparent
                };
                Canvas.SetLeft(searchEllipse, 1);
                Canvas.SetTop(searchEllipse, 1);
                canvas.Children.Add(searchEllipse);
                canvas.Children.Add(new System.Windows.Shapes.Line
                {
                    X1 = 9, Y1 = 9, X2 = 13, Y2 = 13,
                    Stroke = BrushIconGray, StrokeThickness = 1.5
                });
                break;

            case IconKind.Open:
                canvas.Children.Add(new System.Windows.Shapes.Path
                {
                    Data = Geometry.Parse("M3,2 L3,14 L13,14 L13,7 M8,2 L14,2 L14,8 M14,2 L7,9"),
                    Stroke = BrushIconGray, StrokeThickness = 1.3, Fill = Brushes.Transparent
                });
                break;

            case IconKind.Plus:
                canvas.Children.Add(new System.Windows.Shapes.Path
                {
                    Data = Geometry.Parse("M8,3 L8,13 M3,8 L13,8"),
                    Stroke = BrushIconGreen, StrokeThickness = 2
                });
                break;

            case IconKind.Folder:
                canvas.Children.Add(new System.Windows.Shapes.Path
                {
                    Data = Geometry.Parse("M1,4 L1,13 L15,13 L15,6 L7,6 L5.5,4 Z"),
                    Stroke = BrushIconGray, StrokeThickness = 1.2, Fill = BrushFolderFill
                });
                break;
        }

        return canvas;
    }

    // ==============================================================
    // サイドパネルのタブ切替
    // ==============================================================

    // アクティブタブを切り替え、タブヘッダーの見た目とコンテンツの表示を更新
    private void SwitchTab(string tab)
    {
        activeTab = tab;

        // タブヘッダーの見た目を切替
        var tabs = new[] { tabHistory, tabFile, tabSeizure };
        var texts = new[] { tabHistoryText, tabFileText, tabSeizureText };
        var keys = new[] { "history", "file", "seizure" };
        for (int i = 0; i < tabs.Length; i++)
        {
            bool active = keys[i] == tab;
            tabs[i].BorderBrush = active ? BrushAccent : Brushes.Transparent;
            texts[i].Foreground = active ? BrushAccent : BrushTabInactive;
            texts[i].FontWeight = active ? FontWeights.Medium : FontWeights.Normal;
        }

        // タブコンテンツの表示切替
        searchHistoryList.Visibility = tab == "history" ? Visibility.Visible : Visibility.Collapsed;
        fileHistoryList.Visibility   = tab == "file"    ? Visibility.Visible : Visibility.Collapsed;
        seizureLogList.Visibility    = tab == "seizure" ? Visibility.Visible : Visibility.Collapsed;
    }

    // サイドパネルの検索履歴リストを再描画
    private void UpdateSearchHistoryUI()
    {
        searchHistoryList.Items.Clear();
        foreach (var query in searchHistory)
        {
            // 検索アイコン（Canvas Path）
            var icon = new Canvas { Width = 12, Height = 12, Margin = new Thickness(0, 1, 8, 0) };
            var ellipse = new System.Windows.Shapes.Ellipse
            {
                Width = 7, Height = 7,
                Stroke = BrushSecondary, StrokeThickness = 1.2,
                Fill = Brushes.Transparent
            };
            Canvas.SetLeft(ellipse, 0);
            Canvas.SetTop(ellipse, 0);
            icon.Children.Add(ellipse);
            icon.Children.Add(new System.Windows.Shapes.Line
            {
                X1 = 6, Y1 = 6, X2 = 10, Y2 = 10,
                Stroke = BrushSecondary, StrokeThickness = 1.2
            });

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetColumn(icon, 0);
            var queryBlock = new TextBlock
            {
                Text = query,
                FontSize = 12,
                Foreground = BrushAccent,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(queryBlock, 1);
            grid.Children.Add(icon);
            grid.Children.Add(queryBlock);

            searchHistoryList.Items.Add(new ListBoxItem
            {
                Content = grid,
                Padding = new Thickness(10, 15, 10, 15)  // ファイルタブの2行と同等の高さに統一
            });
        }
    }

    // ==============================================================
    // 検索処理（非同期）
    // ==============================================================

    // 宛名番号で .xlsm ファイルを非同期検索し、結果を表示
    // I/O 処理は BackgroundWorker で実行し、UI スレッドのブロックを防止する
    private void ExecuteSearch(string query)
    {
        query = (query ?? "").Trim();
        // 通常モードでは空文字列は無視、お気に入りモードでは空文字列で全件表示
        if (string.IsNullOrEmpty(query) && !isFavoriteMode) return;
        if (searchWorker.IsBusy) return;

        // --- 検索履歴に追加（空文字列は履歴に追加しない） ---
        if (!string.IsNullOrEmpty(query))
        {
            searchHistory.Remove(query);
            searchHistory.Insert(0, query);
            if (searchHistory.Count > config.SearchHistoryMax)
                searchHistory.RemoveAt(config.SearchHistoryMax);
            SaveHistory();
            UpdateSearchHistoryUI();
        }

        // --- UI準備（検索中の操作を制限） ---
        statusLeft.Text = "検索中...";
        statusLeft.Foreground = BrushFooter;
        searchButton.IsEnabled = false;
        searchOverlay.BeginAnimation(UIElement.OpacityProperty, null);  // 前回のフェードアウトをクリア
        searchOverlay.Visibility = Visibility.Visible;

        // --- フェード準備（データ行のみ対象、列ヘッダーは含めない） ---
        if (resultItemsPanel == null)
            resultItemsPanel = FindVisualChild<ItemsPresenter>(resultList);
        if (resultItemsPanel != null)
        {
            resultItemsPanel.Opacity = 0;
            ForceRender();
        }

        // --- 非同期検索開始（検索対象フォルダはモードに応じて切替） ---
        var searchFolders = isFavoriteMode
            ? favoriteFolders.ToArray()
            : config.SearchFolders;
        searchWorker.RunWorkerAsync(new object[] { query, searchFolders });
    }

    // ワーカースレッド: ファイル検索 + ソート（UI スレッドをブロックしない）
    private void SearchWorker_DoWork(object sender, DoWorkEventArgs e)
    {
        var args = (object[])e.Argument;
        var query = (string)args[0];
        var folders = (string[])args[1];
        var results = new List<SearchResultItem>();

        // EnumerateFiles で遅延列挙（GetFiles と異なり全件取得を待たず順次処理可能）
        foreach (var folder in folders)
        {
            if (searchWorker.CancellationPending) { e.Cancel = true; return; }

            try
            {
                var dir = new DirectoryInfo(folder);
                foreach (var file in dir.EnumerateFiles("*.xlsm", SearchOption.AllDirectories))
                {
                    if (searchWorker.CancellationPending) { e.Cancel = true; return; }

                    // BaseName（拡張子なし）に入力文字列を含むか（部分一致・大文字小文字無視）
                    // 空文字列の場合は全件表示（お気に入りモード）
                    var baseName = System.IO.Path.GetFileNameWithoutExtension(file.Name);
                    if (!string.IsNullOrEmpty(query)
                        && baseName.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    var relPath = GetRelativePath(file.FullName);
                    var dirPart = System.IO.Path.GetDirectoryName(relPath) ?? "";

                    results.Add(new SearchResultItem
                    {
                        FileName      = file.Name,
                        FolderPath    = dirPart,
                        LastWriteTime = file.LastWriteTime,
                        LastWrite     = file.LastWriteTime.ToString("MM/dd HH:mm"),
                        SizeBytes     = file.Length,
                        Size          = FormatFileSize(file.Length),
                        FullPath      = file.FullName
                    });
                }
            }
            catch { /* アクセス権エラー等は無視して次のフォルダへ */ }
        }

        // ソートもワーカースレッドで実行（デフォルト: 更新日の降順）
        results.Sort((a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));
        e.Result = results;
    }

    // 検索完了ハンドラ（UI スレッドで実行）
    private void SearchWorker_Completed(object sender, RunWorkerCompletedEventArgs e)
    {
        searchButton.IsEnabled = true;

        // オーバーレイのフェードアウト（100ms → Collapsed）
        var fadeOut = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(100)));
        fadeOut.Completed += delegate { searchOverlay.Visibility = Visibility.Collapsed; };
        searchOverlay.BeginAnimation(UIElement.OpacityProperty, fadeOut);

        if (e.Cancelled)
        {
            // キャンセル後にお気に入りモードの全件再走査が保留中なら実行
            if (pendingFavoriteRefresh && isFavoriteMode)
            {
                pendingFavoriteRefresh = false;
                ExecuteSearch("");
            }
            return;
        }

        if (e.Error != null)
        {
            statusLeft.Text = "検索中にエラーが発生しました";
            statusLeft.Foreground = BrushError;
            return;
        }

        // --- 結果を反映 ---
        currentResults = (List<SearchResultItem>)e.Result;
        currentSortColumn = "LastWriteTime";
        currentSortAscending = false;
        UpdateSortIndicators();
        UpdateResultList();

        // --- フェードイン ---
        if (resultItemsPanel != null)
            resultItemsPanel.BeginAnimation(UIElement.OpacityProperty,
                new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(100))));
    }

    // 検索結果を ListView に反映
    private void UpdateResultList()
    {
        resultList.ItemsSource = null;
        resultList.ItemsSource = currentResults;

        int count = currentResults.Count;
        if (count == 0)
        {
            statusLeft.Text = "該当するファイルが見つかりませんでした";
            statusLeft.Foreground = BrushError;
        }
        else
        {
            statusLeft.Text = "\u2713 " + count + " 件のファイルが見つかりました";
            statusLeft.Foreground = BrushFooter;
        }

        openButton.IsEnabled = false;
        addButton.IsEnabled = false;
    }

    // ==============================================================
    // ソート
    // ==============================================================

    // 列ヘッダークリック時のソート切替
    private void OnColumnHeaderClick(object sender, RoutedEventArgs e)
    {
        var header = e.OriginalSource as GridViewColumnHeader;
        if (header == null || header.Role == GridViewColumnHeaderRole.Padding) return;

        var tag = header.Tag as string;
        if (string.IsNullOrEmpty(tag)) return;

        // 同じ列をクリック → 昇降切替、異なる列 → 昇順で開始
        if (currentSortColumn == tag)
            currentSortAscending = !currentSortAscending;
        else
        {
            currentSortColumn = tag;
            currentSortAscending = true;
        }

        ApplySort();
        UpdateSortIndicators();
        UpdateResultList();
    }

    // 列ヘッダーのソートインジケータ（▲▼）を更新
    private void UpdateSortIndicators()
    {
        var gv = resultList.View as GridView;
        if (gv == null) return;

        foreach (var col in gv.Columns)
        {
            var style = col.HeaderContainerStyle;
            if (style == null) continue;
            var setter = style.Setters.OfType<Setter>()
                .FirstOrDefault(s => s.Property.Name == "Tag");
            if (setter == null) continue;
            var tag = setter.Value as string;
            if (tag == null || !SortColumnNames.ContainsKey(tag)) continue;

            col.Header = (tag == currentSortColumn)
                ? SortColumnNames[tag] + (currentSortAscending ? " \u25B2" : " \u25BC")
                : SortColumnNames[tag];
        }
    }

    // currentResults をインプレースソート（新しい List を生成しない）
    private void ApplySort()
    {
        Comparison<SearchResultItem> comparison;

        switch (currentSortColumn)
        {
            case "FileName":
                comparison = (a, b) => string.Compare(a.FileName, b.FileName, StringComparison.OrdinalIgnoreCase);
                break;
            case "FolderPath":
                comparison = (a, b) => string.Compare(a.FolderPath, b.FolderPath, StringComparison.OrdinalIgnoreCase);
                break;
            case "LastWriteTime":
                comparison = (a, b) => a.LastWriteTime.CompareTo(b.LastWriteTime);
                break;
            case "Size":
                comparison = (a, b) => a.SizeBytes.CompareTo(b.SizeBytes);
                break;
            default:
                return;
        }

        currentResults.Sort(comparison);
        if (!currentSortAscending) currentResults.Reverse();
    }

    // ==============================================================
    // アクション
    // ==============================================================

    // メインパネルまたはサイドパネルで選択中のファイルを開く
    private void OpenSelectedFiles()
    {
        // メインパネルの選択
        var selectedItems = resultList.SelectedItems.Cast<SearchResultItem>().ToList();
        if (selectedItems.Count > 0)
        {
            foreach (var item in selectedItems)
                OpenFile(item.FullPath);
            return;
        }

        // サイドパネルの選択
        var entry = GetSelectedFileHistoryEntry();
        if (entry != null)
            OpenFile(entry.Path);
    }

    // ファイルを OS の既定アプリで開き、ファイル履歴に追加
    private void OpenFile(string fullPath)
    {
        if (!File.Exists(fullPath))
        {
            MessageBox.Show(
                "ファイルが見つかりません:\n" + fullPath,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
            statusLeft.Text = "ファイルを開きました: " + System.IO.Path.GetFileName(fullPath);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "ファイルを開けませんでした:\n" + ex.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        // ファイル履歴に追加
        AddFileHistory(fullPath);
    }

    // メインパネルまたはサイドパネルで選択中のファイルを差押リストに追加
    private void AddSelectedToSeizureList()
    {
        // メインパネルの選択
        var selectedItems = resultList.SelectedItems.Cast<SearchResultItem>().ToList();
        if (selectedItems.Count > 0)
        {
            AddToSeizureList(selectedItems.Select(item => item.FullPath).ToArray());
            return;
        }

        // サイドパネルの選択
        var entry = GetSelectedFileHistoryEntry();
        if (entry != null)
            AddToSeizureList(entry.Path);
    }

    // 選択ファイルの種別を判定し、対応する差押予定一覧 作成ツールを呼び出す
    // ファイル内のセル値で預金/生保を自動判定し、適切なツールに分岐する
    private void AddToSeizureList(params string[] paths)
    {
        if (!config.HasSeizureScript) return;
        if (paths.Length == 0) return;

        // ファイル存在チェック
        var missing = paths.Where(p => !File.Exists(p)).ToArray();
        if (missing.Length > 0)
        {
            MessageBox.Show(
                "ファイルが見つかりません:\n" + string.Join("\n", missing),
                "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // ファイル種別を判定して預金/生保に分類
        var depositPaths = new List<string>();
        var insurancePaths = new List<string>();
        foreach (var path in paths)
        {
            if (IsDepositFile(path)) depositPaths.Add(path);
            else insurancePaths.Add(path);
        }

        // 混在チェック
        if (depositPaths.Count > 0 && insurancePaths.Count > 0)
        {
            MessageBox.Show(
                "預金と生命保険の照会結果が混在しています。\n同じ種類のファイルを選択してください。",
                "確認", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // 呼び出し先の決定
        string scriptPath;
        string toolName;
        if (insurancePaths.Count > 0)
        {
            scriptPath = config.InsuranceScript;
            toolName = "生命保険差押予定一覧 作成ツール";
        }
        else
        {
            scriptPath = config.DepositScript;
            toolName = "預金差押予定一覧 作成ツール";
        }

        if (scriptPath == null)
        {
            MessageBox.Show(
                toolName + "が設定されていません。\nfile_search_config.json を確認してください。",
                "設定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 対象ツールが既に起動中の場合は競合防止のため中断
        var exeName = System.IO.Path.GetFileNameWithoutExtension(scriptPath);
        if (Process.GetProcessesByName(exeName).Length > 0)
        {
            MessageBox.Show(
                toolName + "が起動中です。\n処理が完了してから再度実行してください。",
                "確認", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            // ログパスを構築（file_search.exe と同階層の history/ 配下）
            var logPath = GetSeizureLogPath();
            var logArg = "--log=\"" + logPath + "\"";

            var psi = new ProcessStartInfo
            {
                FileName = scriptPath,
                Arguments = logArg + " " + string.Join(" ", paths.Select(p => "\"" + p + "\"")),
                UseShellExecute = false
            };

            var proc = Process.Start(psi);
            proc.EnableRaisingEvents = true;
            proc.Exited += delegate { Dispatcher.BeginInvoke(new Action(ReloadSeizureLog)); };

            statusLeft.Text = paths.Length == 1
                ? "差押リストに追加中: " + System.IO.Path.GetFileName(paths[0])
                : "差押リストに追加中: " + paths.Length + " 件";
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                toolName + "の呼び出しに失敗しました:\n" + ex.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ==============================================================
    // ファイル種別判定（Open XML）
    // ==============================================================

    private const string NS_SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // .xlsm ファイルを ZIP/Open XML として読み取り、指定セルの値で預金ファイルかを判定する
    // いずれかのシートの depositDetectCell に depositDetectValue が存在すれば預金と判定
    private bool IsDepositFile(string filePath)
    {
        // 判定設定が未構成 → 預金扱い（従来互換）
        if (config.DepositDetectCell == null || config.DepositDetectValue == null)
            return true;

        try
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = Package.Open(fs, FileMode.Open, FileAccess.Read))
            {
                // 表示中シートのファイル名一覧を取得（hidden / veryHidden を除外）
                var visibleSheets = GetVisibleSheetFiles(package);

                // 共有文字列テーブルを読み込む
                var sharedStrings = ReadSharedStrings(package);

                // 表示中シートのみ走査し、指定セルに判定値が存在するか確認
                foreach (var part in package.GetParts())
                {
                    var uri = part.Uri.OriginalString;
                    if (!uri.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)) continue;

                    // workbook.xml で非表示と判定されたシートはスキップ
                    var fileName = uri.Substring(uri.LastIndexOf('/') + 1);
                    if (!visibleSheets.Contains(fileName)) continue;

                    string cellValue = ReadCellValue(part, config.DepositDetectCell, sharedStrings);
                    if (cellValue == config.DepositDetectValue) return true;
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "ファイル種別の判定中にエラーが発生しました。\n預金として処理します。\n\n"
                + System.IO.Path.GetFileName(filePath) + "\n" + ex.GetType().Name + ": " + ex.Message,
                "判定エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return true;
        }
    }

    // workbook.xml と workbook.xml.rels を読み取り、表示中シートのファイル名一覧を返す
    // state 属性がないシート（= visible）のみを対象とする
    private HashSet<string> GetVisibleSheetFiles(Package package)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var wbUri = new Uri("/xl/workbook.xml", UriKind.Relative);
        if (!package.PartExists(wbUri)) return result;

        var wbPart = package.GetPart(wbUri);

        // workbook.xml.rels から rId → ファイル名のマッピングを構築
        var relUri = new Uri("/xl/_rels/workbook.xml.rels", UriKind.Relative);
        var ridToFile = new Dictionary<string, string>();
        if (package.PartExists(relUri))
        {
            using (var stream = package.GetPart(relUri).GetStream(FileMode.Open, FileAccess.Read))
            {
                var doc = new XmlDocument();
                doc.Load(stream);
                foreach (XmlNode node in doc.GetElementsByTagName("Relationship"))
                {
                    var id = node.Attributes["Id"];
                    var target = node.Attributes["Target"];
                    if (id != null && target != null)
                    {
                        // Target="worksheets/sheet1.xml" → "sheet1.xml"
                        var t = target.Value;
                        var fn = t.Substring(t.LastIndexOf('/') + 1);
                        ridToFile[id.Value] = fn;
                    }
                }
            }
        }

        // workbook.xml から表示中シートの rId を取得
        using (var stream = wbPart.GetStream(FileMode.Open, FileAccess.Read))
        {
            var doc = new XmlDocument();
            doc.Load(stream);

            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("s", NS_SPREADSHEET);
            nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            foreach (XmlNode sheet in doc.SelectNodes("//s:sheets/s:sheet", nsmgr))
            {
                // state 属性がない = visible（hidden / veryHidden はスキップ）
                var stateAttr = sheet.Attributes["state"];
                if (stateAttr != null) continue;

                var ridAttr = sheet.Attributes["r:id"];
                if (ridAttr != null && ridToFile.ContainsKey(ridAttr.Value))
                    result.Add(ridToFile[ridAttr.Value]);
            }
        }

        return result;
    }

    // Open XML の共有文字列テーブル（/xl/sharedStrings.xml）を読み込む
    private List<string> ReadSharedStrings(Package package)
    {
        var strings = new List<string>();
        var uri = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
        if (!package.PartExists(uri)) return strings;

        using (var stream = package.GetPart(uri).GetStream(FileMode.Open, FileAccess.Read))
        {
            var doc = new XmlDocument();
            doc.Load(stream);

            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("s", NS_SPREADSHEET);

            foreach (XmlNode si in doc.SelectNodes("//s:si", nsmgr))
            {
                // <si><t>text</t></si> または <si><r><t>text</t></r>...</si>
                var tNodes = si.SelectNodes(".//s:t", nsmgr);
                var sb = new StringBuilder();
                foreach (XmlNode t in tNodes) sb.Append(t.InnerText);
                strings.Add(sb.ToString());
            }
        }
        return strings;
    }

    // シート XML から指定セル（例: "A2"）の値を読み取る
    private string ReadCellValue(PackagePart sheetPart, string cellRef, List<string> sharedStrings)
    {
        using (var stream = sheetPart.GetStream(FileMode.Open, FileAccess.Read))
        {
            var doc = new XmlDocument();
            doc.Load(stream);

            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("s", NS_SPREADSHEET);

            var cellNode = doc.SelectSingleNode(
                "//s:sheetData/s:row/s:c[@r='" + cellRef + "']", nsmgr);
            if (cellNode == null) return null;

            var vNode = cellNode.SelectSingleNode("s:v", nsmgr);
            if (vNode == null) return null;

            // t="s" → 共有文字列テーブルのインデックス参照
            var typeAttr = cellNode.Attributes["t"];
            if (typeAttr != null && typeAttr.Value == "s")
            {
                int idx;
                if (int.TryParse(vNode.InnerText, out idx) && idx < sharedStrings.Count)
                    return sharedStrings[idx];
                return null;
            }

            return vNode.InnerText;
        }
    }

    // --- Win32 API（管理ツールの前面表示用） ---
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    private const int SW_RESTORE = 9;

    // 差押予定一覧 管理ツールを起動する
    // 既に起動中の場合はウィンドウを前面に表示する
    private void LaunchSeizureListManager()
    {
        if (config.SeizureManagerScript == null) return;

        var exeName = System.IO.Path.GetFileNameWithoutExtension(config.SeizureManagerScript);
        var existing = Process.GetProcessesByName(exeName);
        if (existing.Length > 0)
        {
            // 既に起動中 → ウィンドウを前面に表示
            var hwnd = existing[0].MainWindowHandle;
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, SW_RESTORE);
                SetForegroundWindow(hwnd);
            }
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = config.SeizureManagerScript,
                UseShellExecute = false
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "差押予定一覧 管理ツールの起動に失敗しました:\n" + ex.Message,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // 検索結果で右クリックした行のファイルが存在するフォルダを Explorer で開く
    private void OpenSelectedFolder()
    {
        var item = resultList.SelectedItem as SearchResultItem;
        if (item != null) OpenFolder(item.FullPath);
    }

    // ファイルの親フォルダを Explorer で開き、該当ファイルを選択状態にする
    private void OpenFolder(string fullPath)
    {
        if (!File.Exists(fullPath))
        {
            MessageBox.Show(
                "ファイルが見つかりません:\n" + fullPath,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            Process.Start("explorer.exe", "/select," + fullPath);
        }
        catch { /* Explorer の起動失敗は無視 */ }
    }

    // サイドパネルのアクティブタブで選択中の履歴を削除
    private void DeleteSelectedHistory()
    {
        if (activeTab == "history")
        {
            var idx = searchHistoryList.SelectedIndex;
            if (idx < 0 || idx >= searchHistory.Count) return;
            searchHistory.RemoveAt(idx);
            SaveHistory();
            UpdateSearchHistoryUI();
        }
        else
        {
            var idx = fileHistoryList.SelectedIndex;
            if (idx < 0 || idx >= fileHistory.Count) return;
            fileHistory.RemoveAt(idx);
            SaveHistory();
            UpdateFileHistoryUI();
        }
    }

    // サイドパネルの指定タブの履歴を全件削除
    private void ClearAllHistory(string tab)
    {
        if (tab == "history")
        {
            if (searchHistory.Count == 0) return;
            searchHistory.Clear();
            SaveHistory();
            UpdateSearchHistoryUI();
        }
        else
        {
            if (fileHistory.Count == 0) return;
            fileHistory.Clear();
            SaveHistory();
            UpdateFileHistoryUI();
        }
    }

    // ==============================================================
    // 選択状態管理
    // ==============================================================

    // メインパネルまたはサイドパネルに選択がある場合、ボタンを有効化
    private void UpdateButtonState()
    {
        bool hasSelection = resultList.SelectedItem != null
            || fileHistoryList.SelectedIndex >= 0;
        openButton.IsEnabled = hasSelection;
        addButton.IsEnabled = hasSelection && config.HasSeizureScript;
    }

    // サイドパネルで選択中のファイル履歴エントリを返す（未選択時は null）
    private FileHistoryEntry GetSelectedFileHistoryEntry()
    {
        var idx = fileHistoryList.SelectedIndex;
        if (idx < 0 || idx >= fileHistory.Count) return null;
        return fileHistory[idx];
    }

    // ==============================================================
    // 履歴の永続化
    // ==============================================================

    // 履歴ファイルのパスを返す（history\search_{USERNAME}.json）
    private string GetHistoryPath()
    {
        var historyDir = System.IO.Path.Combine(exeDir, "history");
        var username = Environment.UserName;
        return System.IO.Path.Combine(historyDir, "search_" + username + ".json");
    }

    // 起動時に履歴ファイルを読み込む
    private void LoadHistory()
    {
        var path = GetHistoryPath();
        if (!File.Exists(path)) return;

        try
        {
            var json = File.ReadAllText(path, Encoding.UTF8);
            ParseHistoryJson(json);
        }
        catch
        {
            // 読み込み失敗時はセッション内のみで動作（次回保存も試行しない）
            persistHistory = false;
            searchHistory.Clear();
            fileHistory.Clear();
        }
    }

    // 履歴をJSON形式で保存
    private void SaveHistory()
    {
        if (!persistHistory) return;

        try
        {
            var historyDir = System.IO.Path.GetDirectoryName(GetHistoryPath());
            if (!Directory.Exists(historyDir))
                Directory.CreateDirectory(historyDir);

            var sb = new StringBuilder();
            sb.AppendLine("{");

            // searchHistory
            sb.Append("  \"searchHistory\": [");
            for (int i = 0; i < searchHistory.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append("\"" + EscapeJson(searchHistory[i]) + "\"");
            }
            sb.AppendLine("],");

            // fileHistory（path と openedAt のペア）
            sb.AppendLine("  \"fileHistory\": [");
            for (int i = 0; i < fileHistory.Count; i++)
            {
                if (i > 0) sb.AppendLine(",");
                sb.Append("    { \"path\": \"" + EscapeJson(fileHistory[i].Path) + "\", ");
                sb.Append("\"openedAt\": \"" + fileHistory[i].OpenedAt.ToString("o") + "\" }");
            }
            sb.AppendLine();
            sb.AppendLine("  ],");

            // favoriteFolders
            sb.Append("  \"favoriteFolders\": [");
            for (int i = 0; i < favoriteFolders.Count; i++)
            {
                if (i > 0) sb.Append(", ");
                sb.Append("\"" + EscapeJson(favoriteFolders[i]) + "\"");
            }
            sb.AppendLine("],");

            // lastSeenVersion
            sb.AppendLine("  \"lastSeenVersion\": \"" + EscapeJson(lastSeenVersion) + "\"");

            sb.AppendLine("}");

            File.WriteAllText(GetHistoryPath(), sb.ToString(), Encoding.UTF8);
        }
        catch
        {
            // 書き込み失敗時は以降の保存をスキップ（初回のみ）
            if (persistHistory) persistHistory = false;
        }
    }

    // ファイル履歴に追加（重複除去→先頭挿入→上限超過分を削除→即保存）
    private void AddFileHistory(string fullPath)
    {
        fileHistory.RemoveAll(e =>
            string.Equals(e.Path, fullPath, StringComparison.OrdinalIgnoreCase));

        fileHistory.Insert(0, new FileHistoryEntry
        {
            Path = fullPath,
            OpenedAt = DateTime.Now
        });

        if (fileHistory.Count > config.FileHistoryMax)
            fileHistory.RemoveAt(config.FileHistoryMax);

        SaveHistory();
        UpdateFileHistoryUI();
    }

    // サイドパネルの「最近開いたファイル」リストを再描画
    private void UpdateFileHistoryUI()
    {
        fileHistoryList.Items.Clear();
        foreach (var entry in fileHistory)
        {
            var fileName = System.IO.Path.GetFileName(entry.Path);
            var dateStr = entry.OpenedAt.ToString("MM/dd HH:mm");

            var sp = new StackPanel { Margin = new Thickness(0) };

            var nameBlock = new TextBlock
            {
                Text = fileName,
                FontSize = 11,
                Foreground = BrushAccent,
                FontWeight = FontWeights.Medium,
                TextTrimming = TextTrimming.CharacterEllipsis
            };

            var dateBlock = new TextBlock
            {
                Text = dateStr,
                FontSize = 10,
                Foreground = BrushSecondary,
                Margin = new Thickness(0, 2, 0, 0)
            };

            sp.Children.Add(nameBlock);
            sp.Children.Add(dateBlock);

            var item = new ListBoxItem
            {
                Content = sp,
                ToolTip = entry.Path,
                Padding = new Thickness(10, 8, 10, 8)
            };

            fileHistoryList.Items.Add(item);
        }
    }

    // ==============================================================
    // JSON パーサー（手書き・外部ライブラリ不要）
    // LGWAN 環境では NuGet パッケージが使えないため手動パース
    // ==============================================================

    // 履歴JSONを解析して searchHistory / fileHistory / favoriteFolders に格納
    private void ParseHistoryJson(string json)
    {
        searchHistory.Clear();
        fileHistory.Clear();
        favoriteFolders.Clear();

        // --- searchHistory の解析 ---
        var shStart = json.IndexOf("\"searchHistory\"");
        if (shStart >= 0)
        {
            var arrStart = json.IndexOf('[', shStart);
            var arrEnd = json.IndexOf(']', arrStart);
            if (arrStart >= 0 && arrEnd >= 0)
            {
                var arrContent = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
                foreach (var part in SplitJsonStrings(arrContent))
                {
                    var val = part.Trim();
                    if (!string.IsNullOrWhiteSpace(val))
                        searchHistory.Add(val);
                }
            }
        }

        // --- fileHistory の解析 ---
        var fhStart = json.IndexOf("\"fileHistory\"");
        if (fhStart >= 0)
        {
            var arrStart = json.IndexOf('[', fhStart);
            var arrEnd = FindMatchingBracket(json, arrStart);
            if (arrStart >= 0 && arrEnd >= 0)
            {
                var arrContent = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
                int pos = 0;
                while (pos < arrContent.Length)
                {
                    var objStart = arrContent.IndexOf('{', pos);
                    if (objStart < 0) break;
                    var objEnd = arrContent.IndexOf('}', objStart);
                    if (objEnd < 0) break;

                    var objContent = arrContent.Substring(objStart + 1, objEnd - objStart - 1);
                    var path = ExtractJsonStringValue(objContent, "path");
                    var openedAtStr = ExtractJsonStringValue(objContent, "openedAt");

                    if (!string.IsNullOrEmpty(path))
                    {
                        var unescapedPath = path.Replace("\\\\", "\\");
                        DateTime openedAt;
                        if (!DateTime.TryParse(openedAtStr, out openedAt))
                            openedAt = DateTime.MinValue;

                        // 存在しないファイルはスキップ
                        if (File.Exists(unescapedPath))
                        {
                            fileHistory.Add(new FileHistoryEntry
                            {
                                Path = unescapedPath,
                                OpenedAt = openedAt
                            });
                        }
                    }

                    pos = objEnd + 1;
                }
            }
        }

        // --- favoriteFolders の解析 ---
        var ffStart = json.IndexOf("\"favoriteFolders\"");
        if (ffStart >= 0)
        {
            var arrStart = json.IndexOf('[', ffStart);
            var arrEnd = json.IndexOf(']', arrStart);
            if (arrStart >= 0 && arrEnd >= 0)
            {
                var arrContent = json.Substring(arrStart + 1, arrEnd - arrStart - 1);
                foreach (var part in SplitJsonStrings(arrContent))
                {
                    var val = part.Trim().Replace("\\\\", "\\");
                    if (!string.IsNullOrWhiteSpace(val) && Directory.Exists(val)
                        && favoriteFolders.Count < FAVORITE_FOLDERS_MAX)
                        favoriteFolders.Add(val);
                }
            }
        }

        // --- lastSeenVersion の解析 ---
        var lsvVal = ExtractJsonStringValue(json, "lastSeenVersion");
        if (!string.IsNullOrEmpty(lsvVal))
            lastSeenVersion = lsvVal;
    }

    // JSON 文字列配列の中身を分割して返す
    private List<string> SplitJsonStrings(string content)
    {
        var result = new List<string>();
        int pos = 0;
        while (pos < content.Length)
        {
            var qStart = content.IndexOf('"', pos);
            if (qStart < 0) break;
            var qEnd = content.IndexOf('"', qStart + 1);
            if (qEnd < 0) break;
            result.Add(content.Substring(qStart + 1, qEnd - qStart - 1));
            pos = qEnd + 1;
        }
        return result;
    }

    // JSON オブジェクトから指定キーの文字列値を取得
    private string ExtractJsonStringValue(string json, string key)
    {
        var keyIdx = json.IndexOf("\"" + key + "\"");
        if (keyIdx < 0) return null;
        var colonIdx = json.IndexOf(':', keyIdx);
        if (colonIdx < 0) return null;
        var qStart = json.IndexOf('"', colonIdx);
        if (qStart < 0) return null;
        var qEnd = json.IndexOf('"', qStart + 1);
        if (qEnd < 0) return null;
        return json.Substring(qStart + 1, qEnd - qStart - 1);
    }

    // 対応する閉じ括弧 ']' の位置を返す（ネスト対応）
    private int FindMatchingBracket(string json, int openIdx)
    {
        int depth = 0;
        for (int i = openIdx; i < json.Length; i++)
        {
            if (json[i] == '[') depth++;
            else if (json[i] == ']') { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    // JSON 出力用のエスケープ
    private string EscapeJson(string s)
    {
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\n", "\\n").Replace("\r", "\\r").Replace("\t", "\\t");
    }

    // ==============================================================
    // お気に入りフォルダ検索
    // ==============================================================

    // ★ボタン左クリック: 初回はオーバーレイ表示、2回目以降はモード切替
    private void ToggleFavoriteMode()
    {
        if (!isFavoriteMode && favoriteFolders.Count == 0)
        {
            // フォルダ未登録ならオーバーレイを表示して設定を促す
            ShowFavoriteOverlay();
            return;
        }

        isFavoriteMode = !isFavoriteMode;
        ApplyFavoriteModeUI();

        if (isFavoriteMode)
        {
            // ON: テキストクリア → 全件走査開始
            searchBox.Text = "";
            ExecuteSearch("");
        }
        else
        {
            // OFF: テキストクリア → 結果テーブルクリア
            if (searchWorker.IsBusy) searchWorker.CancelAsync();
            searchBox.Text = "";
            currentResults.Clear();
            resultList.ItemsSource = null;
            searchOverlay.BeginAnimation(UIElement.OpacityProperty, null);
            searchOverlay.Visibility = Visibility.Collapsed;
            statusLeft.Text = "";
            statusLeft.Foreground = BrushFooter;
            searchButton.IsEnabled = true;
            openButton.IsEnabled = false;
            addButton.IsEnabled = false;
        }

        searchBox.Focus();
    }

    // ★ボタンとセクションヘッダー・フッターの表示を更新
    private void ApplyFavoriteModeUI()
    {
        // ★ボタンの見た目
        var starBd = FindVisualChild<Border>(starButton);
        if (starBd != null)
        {
            starBd.Background = isFavoriteMode ? BrushStarOn : Brushes.White;
            starBd.BorderBrush = isFavoriteMode ? BrushAccent : BrushBorderLight;
        }
        var starText = FindVisualChild<TextBlock>(starButton);
        if (starText != null)
            starText.Foreground = isFavoriteMode ? BrushAccent : BrushSecondary;

        // セクションヘッダーの★バッジ
        favSectionBadge.Visibility = isFavoriteMode ? Visibility.Visible : Visibility.Collapsed;

        // フッター右
        UpdateFavoriteFooter();
    }

    // フッター右の表示をモードに応じて更新
    private void UpdateFavoriteFooter()
    {
        if (isFavoriteMode)
        {
            statusRight.Text = "★ お気に入りフォルダ（" + favoriteFolders.Count + " 個）";
            statusRight.Foreground = BrushAccent;
            statusRight.ToolTip = string.Join("\n", favoriteFolders);
        }
        else
        {
            statusRight.Foreground = BrushFooter;
            if (config.SearchFolders.Length == 1)
            {
                statusRight.Text = config.SearchFolders[0];
                statusRight.ToolTip = config.SearchFolders[0];
            }
            else
            {
                statusRight.Text = config.SearchFolders.Length + " 個のフォルダを検索中";
                statusRight.ToolTip = string.Join("\n", config.SearchFolders);
            }
        }
    }

    // お気に入り設定オーバーレイを表示（フェードイン）
    private void ShowFavoriteOverlay()
    {
        favFoldersSnapshot = favoriteFolders.ToArray();
        UpdateFavoriteUI();
        favPathInput.Text = "";
        favOverlay.BeginAnimation(UIElement.OpacityProperty, null);
        favOverlay.Opacity = 0;
        favOverlay.Visibility = Visibility.Visible;
        var anim = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150)));
        favOverlay.BeginAnimation(UIElement.OpacityProperty, anim);
        favPathInput.Focus();
    }

    // お気に入り設定オーバーレイを閉じる（フェードアウト）
    private void CloseFavoriteOverlay()
    {
        var anim = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += delegate
        {
            favOverlay.Visibility = Visibility.Collapsed;
            favOverlay.BeginAnimation(UIElement.OpacityProperty, null);
            favOverlay.Opacity = 1;

            if (favoriteFolders.Count > 0)
            {
                if (!isFavoriteMode)
                {
                    // 初回設定完了: モードON → 全件走査
                    isFavoriteMode = true;
                    ApplyFavoriteModeUI();
                    searchBox.Text = "";
                    ExecuteSearch("");
                }
                else
                {
                    // 既にON中の編集: フォルダ構成が変わった場合のみ再走査
                    bool changed = !favoriteFolders.SequenceEqual(favFoldersSnapshot ?? new string[0]);
                    if (changed)
                    {
                        UpdateFavoriteFooter();
                        ExecuteSearch(searchBox.Text);
                    }
                }
            }
            searchBox.Focus();
        };
        favOverlay.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    // お気に入りフォルダを追加（パス入力欄から）
    private void AddFavoriteFolder()
    {
        var path = (favPathInput.Text ?? "").Trim();
        if (string.IsNullOrEmpty(path)) return;

        if (!Directory.Exists(path))
        {
            MessageBox.Show("フォルダが見つかりません:\n" + path,
                "エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // 重複チェック
        if (favoriteFolders.Any(f => string.Equals(f, path, StringComparison.OrdinalIgnoreCase)))
        {
            MessageBox.Show("このフォルダは既に登録されています。",
                "確認", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        if (favoriteFolders.Count >= FAVORITE_FOLDERS_MAX)
        {
            MessageBox.Show("お気に入りフォルダは最大 " + FAVORITE_FOLDERS_MAX + " 個までです。",
                "確認", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        favoriteFolders.Add(path);
        SaveHistory();
        UpdateFavoriteUI();
        favPathInput.Text = "";
        favPathInput.Focus();
    }

    // お気に入りフォルダを削除（インデックス指定）
    private void RemoveFavoriteFolder(int index)
    {
        if (index < 0 || index >= favoriteFolders.Count) return;
        favoriteFolders.RemoveAt(index);
        SaveHistory();
        UpdateFavoriteUI();

        // 全削除された場合はモードOFF
        if (favoriteFolders.Count == 0 && isFavoriteMode)
        {
            isFavoriteMode = false;
            ApplyFavoriteModeUI();
            currentResults.Clear();
            resultList.ItemsSource = null;
            statusLeft.Text = "";
            statusLeft.Foreground = BrushFooter;
            openButton.IsEnabled = false;
            addButton.IsEnabled = false;
        }
    }

    // お気に入りフォルダスロットの表示を更新
    private void UpdateFavoriteUI()
    {
        favSlotPanel.Children.Clear();

        // 登録済みスロット
        for (int i = 0; i < favoriteFolders.Count; i++)
        {
            var idx = i;  // クロージャ用
            var border = new Border
            {
                BorderBrush = BrushBorderNormal,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8, 7, 8, 7),
                Margin = new Thickness(0, 0, 0, 6)
            };

            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) });

            // フォルダアイコン（Canvas+Path、MakeIcon と同一方式）
            var icon = MakeIcon(IconKind.Folder);
            Grid.SetColumn((UIElement)icon, 0);
            grid.Children.Add((UIElement)icon);

            // パステキスト
            var pathText = new TextBlock
            {
                Text = favoriteFolders[i],
                FontSize = 11, FontFamily = new FontFamily("Consolas"),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                Margin = new Thickness(4, 0, 4, 0),
                ToolTip = favoriteFolders[i]
            };
            Grid.SetColumn(pathText, 1);
            grid.Children.Add(pathText);

            // 削除ボタン（ControlTemplate でデフォルトテーマを除去、ホバー時に赤）
            var delBtn = new Button { Cursor = Cursors.Hand, Focusable = false };
            var delTemplate = new ControlTemplate(typeof(Button));
            var delBorder = new FrameworkElementFactory(typeof(Border));
            delBorder.SetValue(Border.BackgroundProperty, Brushes.Transparent);
            delBorder.SetValue(Border.WidthProperty, 20.0);
            delBorder.SetValue(Border.HeightProperty, 20.0);
            delBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));
            delBorder.Name = "delBg";
            var delText = new FrameworkElementFactory(typeof(TextBlock));
            delText.SetValue(TextBlock.TextProperty, "✕");
            delText.SetValue(TextBlock.FontSizeProperty, 10.0);
            delText.SetValue(TextBlock.ForegroundProperty, BrushSecondary);
            delText.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            delText.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            delText.Name = "delTxt";
            delBorder.AppendChild(delText);
            delTemplate.VisualTree = delBorder;
            var hoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, BrushHoverBg, "delBg"));
            delTemplate.Triggers.Add(hoverTrigger);
            delBtn.Template = delTemplate;
            delBtn.Click += delegate { RemoveFavoriteFolder(idx); };
            Grid.SetColumn(delBtn, 2);
            grid.Children.Add(delBtn);

            border.Child = grid;
            favSlotPanel.Children.Add(border);
        }

        // 空スロット
        for (int i = favoriteFolders.Count; i < FAVORITE_FOLDERS_MAX; i++)
        {
            var border = new Border
            {
                BorderBrush = BrushBorderLight,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(8, 7, 8, 7),
                Margin = new Thickness(0, 0, 0, 6),
                Background = BrushSlotEmpty
            };

            var sp = new StackPanel { Orientation = Orientation.Horizontal };
            sp.Children.Add(new TextBlock
            {
                Text = (i + 1).ToString(),
                FontSize = 11, Foreground = BrushSecondary,
                Width = 16, VerticalAlignment = VerticalAlignment.Center
            });
            sp.Children.Add(new TextBlock
            {
                Text = "未登録",
                FontSize = 11, Foreground = BrushSecondary,
                VerticalAlignment = VerticalAlignment.Center
            });

            border.Child = sp;
            favSlotPanel.Children.Add(border);
        }
    }

    // ==============================================================
    // 差押登録ログ
    // ==============================================================

    // 差押登録ログファイルのパスを返す（history/seizure_log_{USERNAME}.json）
    private string GetSeizureLogPath()
    {
        var histDir = System.IO.Path.Combine(exeDir, "history");
        return System.IO.Path.Combine(histDir, "seizure_log_" + Environment.UserName + ".json");
    }

    // 差押登録ログを読み込み seizureLog に格納
    private void LoadSeizureLog()
    {
        seizureLog.Clear();
        try
        {
            var logPath = GetSeizureLogPath();
            if (!File.Exists(logPath)) return;

            var json = File.ReadAllText(logPath, Encoding.UTF8);
            var entriesStart = json.IndexOf("\"entries\"");
            if (entriesStart < 0) return;
            var arrStart = json.IndexOf('[', entriesStart);
            var arrEnd = FindMatchingBracket(json, arrStart);
            if (arrStart < 0 || arrEnd < 0) return;
            var arrContent = json.Substring(arrStart + 1, arrEnd - arrStart - 1);

            int pos = 0;
            while (pos < arrContent.Length && seizureLog.Count < config.SeizureLogMax)
            {
                var objStart = arrContent.IndexOf('{', pos);
                if (objStart < 0) break;
                var objEnd = arrContent.IndexOf('}', objStart);
                if (objEnd < 0) break;
                var objContent = arrContent.Substring(objStart + 1, objEnd - objStart - 1);

                var entry = new SeizureLogEntry
                {
                    AddressNumber   = ExtractJsonStringValue(objContent, "addressNumber") ?? "",
                    Name            = ExtractJsonStringValue(objContent, "name") ?? "",
                    InstitutionName = ExtractJsonStringValue(objContent, "institutionName") ?? "",
                    ExecutionDate   = ExtractJsonStringValue(objContent, "executionDate") ?? "",
                    DocumentNumber  = ExtractJsonStringValue(objContent, "documentNumber") ?? ""
                };

                var regStr = ExtractJsonStringValue(objContent, "registeredAt");
                DateTime regDt;
                entry.RegisteredAt = DateTime.TryParse(regStr, out regDt) ? regDt : DateTime.MinValue;

                seizureLog.Add(entry);
                pos = objEnd + 1;
            }
        }
        catch { /* 読み込み失敗時は空のまま */ }
    }

    // 子プロセス終了後にログを再読み込みしてUIを更新
    private void ReloadSeizureLog()
    {
        LoadSeizureLog();
        UpdateSeizureLogUI();
    }

    // サイドパネルの差押登録タブを再描画
    private void UpdateSeizureLogUI()
    {
        seizureLogList.Items.Clear();
        foreach (var entry in seizureLog)
        {
            var sp = new StackPanel { Margin = new Thickness(0, 0, 0, 0) };

            // 1行目: 宛名番号 + 氏名
            var line1 = new Grid();
            line1.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            line1.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var addrBlock = new TextBlock
            {
                Text = entry.AddressNumber,
                FontSize = 11,
                Foreground = BrushAccent,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 5, 0)
            };
            Grid.SetColumn(addrBlock, 0);
            var nameBlock = new TextBlock
            {
                Text = entry.Name,
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis
            };
            Grid.SetColumn(nameBlock, 1);
            line1.Children.Add(addrBlock);
            line1.Children.Add(nameBlock);
            sp.Children.Add(line1);

            // 2行目: 登録日時
            sp.Children.Add(new TextBlock
            {
                Text = entry.RegisteredAt.ToString("MM/dd HH:mm"),
                FontSize = 10,
                Foreground = BrushSecondary,
                Margin = new Thickness(0, 2, 0, 0)
            });

            // ToolTip: 金融機関名・文書番号・執行日
            var toolTip = "金融機関: " + entry.InstitutionName
                + "\n文書番号: " + entry.DocumentNumber
                + "\n執行日: " + entry.ExecutionDate;

            var item = new ListBoxItem
            {
                Content = sp,
                Padding = new Thickness(10, 8, 10, 8),
                BorderBrush = BrushHoverBg,
                BorderThickness = new Thickness(0, 0, 0, 1),
                ToolTip = toolTip
            };
            seizureLogList.Items.Add(item);
        }
    }

    // ==============================================================
    // アップデート告知
    // ==============================================================

    // 起動時にアップデート告知の表示要否をチェック
    private void CheckUpdateNotice()
    {
        try
        {
            var noticePath = System.IO.Path.Combine(exeDir, "update_notice.json");
            if (!File.Exists(noticePath)) return;

            var notice = LoadUpdateNotice(noticePath);
            if (notice == null || string.IsNullOrEmpty(notice.Version)) return;
            if (notice.Version == lastSeenVersion) return;
            if (notice.Features.Count == 0) return;

            pendingNoticeVersion = notice.Version;
            BuildNoticeUI(notice);
            ShowNoticeOverlay();
        }
        catch { /* 告知読み込み失敗時は通知をスキップ */ }
    }

    // update_notice.json を手動パースして UpdateNotice を返す
    private UpdateNotice LoadUpdateNotice(string path)
    {
        var json = File.ReadAllText(path, Encoding.UTF8);
        var notice = new UpdateNotice();

        // version
        notice.Version = ExtractJsonStringValue(json, "version");

        // features 配列の範囲を取得
        var featuresStart = json.IndexOf("\"features\"");
        if (featuresStart < 0) return notice;
        var featArrStart = json.IndexOf('[', featuresStart);
        var featArrEnd = FindMatchingBracket(json, featArrStart);
        if (featArrStart < 0 || featArrEnd < 0) return notice;
        var featArrContent = json.Substring(featArrStart + 1, featArrEnd - featArrStart - 1);

        // 各 feature オブジェクトを解析
        int pos = 0;
        while (pos < featArrContent.Length)
        {
            var objStart = featArrContent.IndexOf('{', pos);
            if (objStart < 0) break;
            var objEnd = FindMatchingBrace(featArrContent, objStart);
            if (objEnd < 0) break;
            var objContent = featArrContent.Substring(objStart, objEnd - objStart + 1);

            var feature = new NoticeFeature();
            feature.Title = ExtractJsonStringValue(objContent, "title") ?? "";
            feature.Description = ExtractJsonStringValue(objContent, "description") ?? "";

            // sections 配列
            var secStart = objContent.IndexOf("\"sections\"");
            if (secStart >= 0)
            {
                var secArrStart = objContent.IndexOf('[', secStart);
                var secArrEnd = FindMatchingBracket(objContent, secArrStart);
                if (secArrStart >= 0 && secArrEnd >= 0)
                {
                    var secContent = objContent.Substring(secArrStart + 1, secArrEnd - secArrStart - 1);
                    int sp = 0;
                    while (sp < secContent.Length)
                    {
                        var sObjStart = secContent.IndexOf('{', sp);
                        if (sObjStart < 0) break;
                        var sObjEnd = secContent.IndexOf('}', sObjStart);
                        if (sObjEnd < 0) break;
                        var sObj = secContent.Substring(sObjStart + 1, sObjEnd - sObjStart - 1);

                        var section = new NoticeSection();
                        section.Heading = ExtractJsonStringValue(sObj, "heading") ?? "";

                        // items 配列
                        var itemsStart = sObj.IndexOf("\"items\"");
                        if (itemsStart >= 0)
                        {
                            var iArrStart = sObj.IndexOf('[', itemsStart);
                            var iArrEnd = sObj.IndexOf(']', iArrStart);
                            if (iArrStart >= 0 && iArrEnd >= 0)
                            {
                                var iContent = sObj.Substring(iArrStart + 1, iArrEnd - iArrStart - 1);
                                section.Items = SplitJsonStrings(iContent);
                            }
                        }

                        feature.Sections.Add(section);
                        sp = sObjEnd + 1;
                    }
                }
            }

            notice.Features.Add(feature);
            pos = objEnd + 1;
        }

        return notice;
    }

    // 対応する閉じ中括弧 '}' の位置を返す（ネスト対応）
    private int FindMatchingBrace(string json, int openIdx)
    {
        int depth = 0;
        for (int i = openIdx; i < json.Length; i++)
        {
            if (json[i] == '{') depth++;
            else if (json[i] == '}') { depth--; if (depth == 0) return i; }
        }
        return -1;
    }

    // UpdateNotice の内容で告知オーバーレイを動的構築
    private void BuildNoticeUI(UpdateNotice notice)
    {
        noticeContent.Children.Clear();

        // バージョンバッジ
        var versionBadge = new TextBlock
        {
            Text = "ver " + notice.Version,
            FontSize = 11,
            Foreground = BrushAccent,
            Padding = new Thickness(10, 2, 10, 5)
        };
        var badgeBorder = new Border
        {
            Background = BrushStarOn,
            CornerRadius = new CornerRadius(10),
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 0, 0, 16),
            Child = versionBadge
        };
        noticeContent.Children.Add(badgeBorder);

        // 区切り線
        noticeContent.Children.Add(new Border
        {
            BorderBrush = BrushBorderNormal,
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(0, 16, 0, 0)
        });

        // 各 feature ブロック
        for (int fi = 0; fi < notice.Features.Count; fi++)
        {
            var feature = notice.Features[fi];
            if (fi > 0)
                noticeContent.Children.Add(new Border { Height = 12 });

            // タイトル
            noticeContent.Children.Add(new TextBlock
            {
                Text = feature.Title,
                FontSize = 13, FontWeight = FontWeights.Medium,
                Foreground = BrushAccent,
                Margin = new Thickness(0, 0, 0, 8)
            });

            // 説明文
            if (!string.IsNullOrEmpty(feature.Description))
            {
                noticeContent.Children.Add(new TextBlock
                {
                    Text = feature.Description,
                    FontSize = 12, Foreground = BrushFooter,
                    TextWrapping = TextWrapping.Wrap,
                    LineHeight = 20,
                    Margin = new Thickness(0, 0, 0, 12)
                });
            }

            // セクション
            for (int si = 0; si < feature.Sections.Count; si++)
            {
                var section = feature.Sections[si];
                noticeContent.Children.Add(new TextBlock
                {
                    Text = section.Heading,
                    FontSize = 12, FontWeight = FontWeights.Medium,
                    Margin = new Thickness(0, 0, 0, 4)
                });

                foreach (var item in section.Items)
                {
                    noticeContent.Children.Add(new TextBlock
                    {
                        Text = item,
                        FontSize = 12, Foreground = BrushFooter,
                        TextWrapping = TextWrapping.Wrap,
                        LineHeight = 20,
                        Margin = new Thickness(0, 0, 0, 2)
                    });
                }

                // セクション間のスペーサー（最後のセクションには追加しない）
                if (si < feature.Sections.Count - 1)
                    noticeContent.Children.Add(new Border { Height = 8 });
            }
        }
    }

    // アップデート告知オーバーレイを表示（フェードイン）
    private void ShowNoticeOverlay()
    {
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty, null);
        noticeOverlay.Opacity = 0;
        noticeOverlay.Visibility = Visibility.Visible;
        var anim = new DoubleAnimation(0, 1, new Duration(TimeSpan.FromMilliseconds(150)));
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    // アップデート告知オーバーレイを閉じる（フェードアウト → lastSeenVersion 保存）
    private void CloseNoticeOverlay()
    {
        if (!string.IsNullOrEmpty(pendingNoticeVersion))
            lastSeenVersion = pendingNoticeVersion;
        SaveHistory();

        var anim = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromMilliseconds(150)));
        anim.Completed += delegate
        {
            noticeOverlay.Visibility = Visibility.Collapsed;
            noticeOverlay.BeginAnimation(UIElement.OpacityProperty, null);
            noticeOverlay.Opacity = 1;
            searchBox.Focus();
        };
        noticeOverlay.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    // ==============================================================
    // ヘルパー
    // ==============================================================

    // ファイルのフルパスから、該当する検索フォルダを基準にした相対パスを返す
    private string GetRelativePath(string fullPath)
    {
        // searchFolder とお気に入りフォルダの両方を対象に検索
        var allFolders = config.SearchFolders.Concat(favoriteFolders);
        foreach (var folder in allFolders)
        {
            var normalized = folder.TrimEnd('\\');
            if (fullPath.StartsWith(normalized + "\\", StringComparison.OrdinalIgnoreCase))
                return fullPath.Substring(normalized.Length + 1);
        }
        return System.IO.Path.GetFileName(fullPath);
    }

    // ファイルサイズを読みやすい文字列にフォーマット
    private string FormatFileSize(long bytes)
    {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return (bytes / 1024) + " KB";
        return (bytes / (1024.0 * 1024.0)).ToString("N1") + " MB";
    }

    // ==============================================================
    // XAML 定義
    // ==============================================================

    // ウィンドウの XAML をインライン文字列で定義し、XamlReader.Load で読み込む
    // .NET Framework 4.0 の csc.exe では XAML ファイルの埋め込み（BAML）が使えないため
    private Window BuildWindow()
    {
        // 差押ツールの有無に応じてボタンの XAML を動的生成
        string addButtonXaml = config.HasSeizureScript
            ? @"<Button x:Name='AddButton' Grid.Column='2'
                        Style='{StaticResource AB}' IsEnabled='False'>
                    <StackPanel Orientation='Horizontal'>
                        <Canvas Width='14' Height='14' Margin='0,0,4,0'>
                            <Path Data='M7,2 L7,12 M2,7 L12,7'
                                  Stroke='White' StrokeThickness='2'/>
                        </Canvas>
                        <TextBlock Text='差押リストに追加' VerticalAlignment='Center'/>
                    </StackPanel>
                </Button>"
            : @"<Button x:Name='AddButton' Grid.Column='2'
                        Visibility='Collapsed' Style='{StaticResource AB}'/>";

        // seizureManagerScript の有無に応じて管理ツール起動ボタンを動的生成
        string managerButtonXaml = config.SeizureManagerScript != null
            ? @"<Button x:Name='ManagerButton' Grid.Column='4'
                        Style='{StaticResource MB}' Padding='0' Width='34' Height='34'
                        ToolTip='差押予定一覧管理ツール'>
                    <Canvas Width='26' Height='26'>
                        <Path Data='M11.0,5.7 L11.8,3.2 L14.2,3.2 L15.0,5.7 L16.7,6.4 L19.1,5.2 L20.8,6.9 L19.6,9.3 L20.3,11.0 L22.8,11.8 L22.8,14.2 L20.3,15.0 L19.6,16.7 L20.8,19.1 L19.1,20.8 L16.7,19.6 L15.0,20.3 L14.2,22.8 L11.8,22.8 L11.0,20.3 L9.3,19.6 L6.9,20.8 L5.2,19.1 L6.4,16.7 L5.7,15.0 L3.2,14.2 L3.2,11.8 L5.7,11.0 L6.4,9.3 L5.2,6.9 L6.9,5.2 L9.3,6.4 Z'
                              Fill='White'/>
                        <Ellipse Canvas.Left='10.1' Canvas.Top='10.1' Width='5.8' Height='5.8'
                                 Fill='#546E7A'/>
                    </Canvas>
                </Button>"
            : @"<Button x:Name='ManagerButton' Grid.Column='4'
                        Visibility='Collapsed' Style='{StaticResource MB}'/>";

        // 各ボタンの設定有無に応じて列定義を構築
        string colGap1 = config.HasSeizureScript
            ? "<ColumnDefinition Width='8'/>" : "<ColumnDefinition Width='0'/>";
        string colAdd = config.HasSeizureScript
            ? "<ColumnDefinition Width='*'/>" : "<ColumnDefinition Width='0'/>";
        string colGap2 = config.SeizureManagerScript != null
            ? "<ColumnDefinition Width='8'/>" : "<ColumnDefinition Width='0'/>";
        string colManager = config.SeizureManagerScript != null
            ? "<ColumnDefinition Width='Auto'/>" : "<ColumnDefinition Width='0'/>";
        string buttonColumns = "<ColumnDefinition Width='*'/>" + colGap1 + colAdd + colGap2 + colManager;

        string xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    Title='預貯金照会結果 検索ツール'
    Width='1000' Height='700' MinWidth='900' MinHeight='520'
    WindowStartupLocation='CenterScreen'
    Background='#F9F9F9' FontFamily='Meiryo UI'
    UseLayoutRounding='True'
    SnapsToDevicePixels='True'
    TextOptions.TextFormattingMode='Display'
    TextOptions.TextRenderingMode='ClearType'>

    <Window.Resources>
        <!-- アクセントボタン（プライマリ: 青背景） -->
        <Style x:Key='AB' TargetType='Button'>
            <Setter Property='Background' Value='#005FB8'/><Setter Property='Foreground' Value='White'/>
            <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
            <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderThickness' Value='0'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='{TemplateBinding Background}'
                            CornerRadius='4' Padding='{TemplateBinding Padding}'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#004FA0'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Background' Value='#CCC'/>
                            <Setter Property='Foreground' Value='#999'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- ゴーストボタン（セカンダリ: 白背景＋ボーダー） -->
        <Style x:Key='GB' TargetType='Button'>
            <Setter Property='Background' Value='White'/><Setter Property='Foreground' Value='#555'/>
            <Setter Property='FontSize' Value='12'/><Setter Property='Padding' Value='16,8'/>
            <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderBrush' Value='#D0D0D0'/><Setter Property='BorderThickness' Value='1'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='{TemplateBinding Background}'
                            BorderBrush='{TemplateBinding BorderBrush}'
                            BorderThickness='{TemplateBinding BorderThickness}'
                            CornerRadius='4' Padding='{TemplateBinding Padding}'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#EBF1F8'/></Trigger>
                        <Trigger Property='IsEnabled' Value='False'>
                            <Setter TargetName='bd' Property='Background' Value='#F5F5F5'/>
                            <Setter Property='Foreground' Value='#CCC'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 管理ツールボタン（スレート背景 + 白アイコン） -->
        <Style x:Key='MB' TargetType='Button'>
            <Setter Property='Background' Value='#546E7A'/><Setter Property='Foreground' Value='White'/>
            <Setter Property='Cursor' Value='Hand'/><Setter Property='BorderThickness' Value='0'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='Button'>
                    <Border x:Name='bd' Background='{TemplateBinding Background}'
                            CornerRadius='4' Padding='{TemplateBinding Padding}'>
                        <ContentPresenter HorizontalAlignment='Center' VerticalAlignment='Center'/></Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='bd' Property='Background' Value='#455A64'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- GridViewColumnHeader: フラットデザイン（中央揃え） -->
        <Style TargetType='GridViewColumnHeader'>
            <Setter Property='Background' Value='#F5F7FA'/>
            <Setter Property='Foreground' Value='#777'/>
            <Setter Property='FontSize' Value='11'/>
            <Setter Property='Padding' Value='8,7'/>
            <Setter Property='HorizontalContentAlignment' Value='Center'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='GridViewColumnHeader'>
                    <Border Background='{TemplateBinding Background}'
                            BorderBrush='#E8E8E8' BorderThickness='0,0,0,1'
                            Padding='{TemplateBinding Padding}'>
                        <ContentPresenter HorizontalAlignment='{TemplateBinding HorizontalContentAlignment}'
                                          VerticalAlignment='Center'/></Border>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- 検索結果 ListViewItem: アクセントバー＋ホバー/選択 -->
        <Style x:Key='ResultItem' TargetType='ListViewItem'>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Foreground' Value='#333'/>
            <Setter Property='HorizontalContentAlignment' Value='Stretch'/>
            <Setter Property='Padding' Value='0'/><Setter Property='Margin' Value='0'/>
            <Setter Property='Template'><Setter.Value>
                <ControlTemplate TargetType='ListViewItem'>
                    <Grid>
                        <Border x:Name='rowBd' Background='White'
                                BorderBrush='#F0F0F0' BorderThickness='0,0,0,1'>
                            <GridViewRowPresenter VerticalAlignment='Center'
                                Margin='0,7,0,7'/></Border>
                        <Border x:Name='accent' HorizontalAlignment='Left'
                                Width='3' Background='Transparent'/></Grid>
                    <ControlTemplate.Triggers>
                        <Trigger Property='IsMouseOver' Value='True'>
                            <Setter TargetName='rowBd' Property='Background' Value='#F0F4F8'/></Trigger>
                        <Trigger Property='IsSelected' Value='True'>
                            <Setter TargetName='accent' Property='Background' Value='#005FB8'/>
                            <Setter TargetName='rowBd' Property='Background' Value='#E8F0FE'/></Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value></Setter>
        </Style>

        <!-- サイドパネル ListBoxItem -->
        <Style x:Key='HistoryItem' TargetType='ListBoxItem'>
            <Setter Property='HorizontalContentAlignment' Value='Stretch'/>
            <Setter Property='Padding' Value='0'/>
            <Setter Property='Cursor' Value='Hand'/>
            <Setter Property='Template'>
                <Setter.Value>
                    <ControlTemplate TargetType='ListBoxItem'>
                        <Border x:Name='bg' Padding='{TemplateBinding Padding}'
                                Background='Transparent' BorderBrush='#F0F0F0'
                                BorderThickness='0,0,0,1'>
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property='IsMouseOver' Value='True'>
                                <Setter TargetName='bg' Property='Background' Value='#F0F4F8'/>
                            </Trigger>
                            <Trigger Property='IsSelected' Value='True'>
                                <Setter TargetName='bg' Property='Background' Value='#E8F0FE'/>
                                <Setter TargetName='bg' Property='BorderBrush' Value='#E8F0FE'/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid>
    <DockPanel>
        <!-- ============ ヘッダー ============ -->
        <Border DockPanel.Dock='Top' Background='#005FB8' Padding='18,10'>
            <TextBlock Text='預貯金照会結果 検索ツール'
                       FontSize='13' FontWeight='Medium' Foreground='White'/>
        </Border>

        <!-- ============ フッター ============ -->
        <Border DockPanel.Dock='Bottom' Background='#F0F0F0'
                BorderBrush='#E0E0E0' BorderThickness='0,1,0,0' Padding='18,4'>
            <DockPanel>
                <TextBlock x:Name='StatusRight' DockPanel.Dock='Right'
                           FontSize='11' Foreground='#666'/>
                <TextBlock x:Name='StatusLeft' FontSize='11' Foreground='#666'/>
            </DockPanel>
        </Border>

        <!-- ============ メインエリア ============ -->
        <Grid Margin='18,14,18,12'>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width='*'/>
                <ColumnDefinition Width='12'/>
                <ColumnDefinition Width='210'/>
            </Grid.ColumnDefinitions>

            <!-- === 左カラム === -->
            <DockPanel Grid.Column='0'>

                <!-- 検索バー -->
                <Grid DockPanel.Dock='Top' Margin='0,0,0,12'>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width='*'/>
                        <ColumnDefinition Width='8'/>
                        <ColumnDefinition Width='Auto'/>
                        <ColumnDefinition Width='8'/>
                        <ColumnDefinition Width='Auto'/>
                    </Grid.ColumnDefinitions>

                    <Border Grid.Column='0' Background='White'
                            BorderBrush='#D0D0D0' BorderThickness='1' CornerRadius='4'>
                        <Grid>
                            <TextBlock x:Name='SearchPlaceholder'
                                       Text='宛名番号を入力...' FontSize='13'
                                       Foreground='#999' VerticalAlignment='Center'
                                       Margin='32,0,0,0' IsHitTestVisible='False'/>
                            <Canvas Width='14' Height='14' VerticalAlignment='Center'
                                   HorizontalAlignment='Left'
                                   Margin='10,0,0,0' IsHitTestVisible='False'>
                                <Ellipse Canvas.Left='1' Canvas.Top='1' Width='8' Height='8'
                                         Stroke='#999' StrokeThickness='1.5' Fill='Transparent'/>
                                <Line X1='8' Y1='8' X2='12' Y2='12'
                                      Stroke='#999' StrokeThickness='1.5'/>
                            </Canvas>
                            <TextBox x:Name='SearchBox' FontSize='13' FontFamily='Consolas'
                                     Padding='28,8,28,8' BorderThickness='0' Background='Transparent'
                                     VerticalContentAlignment='Center'/>
                            <Button x:Name='ClearButton' HorizontalAlignment='Right'
                                    VerticalAlignment='Center' Margin='0,0,6,0'
                                    Visibility='Collapsed' Cursor='Hand'
                                    Focusable='False'>
                                <Button.Template>
                                    <ControlTemplate TargetType='Button'>
                                        <Border x:Name='bg' Width='20' Height='20'
                                                CornerRadius='10' Background='#E0E0E0'>
                                            <Canvas Width='10' Height='10'
                                                    HorizontalAlignment='Center' VerticalAlignment='Center'>
                                                <Line x:Name='l1' X1='1' Y1='1' X2='9' Y2='9'
                                                      Stroke='#777' StrokeThickness='1.5'
                                                      StrokeStartLineCap='Round' StrokeEndLineCap='Round'/>
                                                <Line x:Name='l2' X1='9' Y1='1' X2='1' Y2='9'
                                                      Stroke='#777' StrokeThickness='1.5'
                                                      StrokeStartLineCap='Round' StrokeEndLineCap='Round'/>
                                            </Canvas>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property='IsMouseOver' Value='True'>
                                                <Setter TargetName='bg' Property='Background' Value='#CCC'/>
                                                <Setter TargetName='l1' Property='Stroke' Value='#555'/>
                                                <Setter TargetName='l2' Property='Stroke' Value='#555'/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Button.Template>
                            </Button>
                        </Grid>
                    </Border>


                    <Button x:Name='SearchButton' Grid.Column='2'
                            Style='{StaticResource AB}'>
                        <StackPanel Orientation='Horizontal' Margin='-2,0,2,0'>
                            <Canvas Width='14' Height='14' Margin='0,0,4,0'>
                                <Ellipse Canvas.Left='1' Canvas.Top='1' Width='8' Height='8'
                                         Stroke='White' StrokeThickness='1.5' Fill='Transparent'/>
                                <Line X1='8' Y1='8' X2='12' Y2='12'
                                      Stroke='White' StrokeThickness='1.5'/>
                            </Canvas>
                            <TextBlock Text='検索' VerticalAlignment='Center'/>
                        </StackPanel>
                    </Button>

                    <!-- ★お気に入りフォルダ検索ボタン -->
                    <Button x:Name='StarButton' Grid.Column='4'
                            Cursor='Hand' ToolTip='お気に入りフォルダ検索'>
                        <Button.Template>
                            <ControlTemplate TargetType='Button'>
                                <Border x:Name='starBd' Background='White'
                                        BorderBrush='#D0D0D0' BorderThickness='1'
                                        CornerRadius='4' Width='32' Height='32'>
                                    <TextBlock Text='★' FontSize='15'
                                               Foreground='#999'
                                               HorizontalAlignment='Center'
                                               VerticalAlignment='Center'
                                               Margin='0,-1,0,0'/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property='IsMouseOver' Value='True'>
                                        <Setter TargetName='starBd' Property='Background' Value='#F0F4F8'/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                </Grid>

                <!-- アクションボタン -->
                <Grid DockPanel.Dock='Bottom' Margin='0,12,0,0'>
                    <Grid.ColumnDefinitions>
                        " + buttonColumns + @"
                    </Grid.ColumnDefinitions>
                    <Button x:Name='OpenButton' Grid.Column='0'
                            Style='{StaticResource GB}' IsEnabled='False'>
                        <StackPanel Orientation='Horizontal'>
                            <Canvas Width='14' Height='14' Margin='0,0,4,0'>
                                <Path Data='M3,2 L3,12 L11,12 L11,6 M7,2 L12,2 L12,7 M12,2 L6,8'
                                      Stroke='#555' StrokeThickness='1.3' Fill='Transparent'/>
                            </Canvas>
                            <TextBlock Text='ファイルを開く' VerticalAlignment='Center'/>
                        </StackPanel>
                    </Button>
                    " + addButtonXaml + @"
                    " + managerButtonXaml + @"
                </Grid>

                <!-- 検索結果パネル -->
                <Border Background='White' BorderBrush='#E0E0E0' BorderThickness='1'
                        CornerRadius='4'>
                    <DockPanel>
                        <Border DockPanel.Dock='Top' Padding='14,8'
                                Background='#FAFAFA'
                                BorderBrush='#E0E0E0' BorderThickness='0,0,0,1'>
                            <DockPanel>
                                <TextBlock Text='&#x2261; 検索結果'
                                           FontSize='13' Foreground='#005FB8' FontWeight='Medium'/>
                                <TextBlock x:Name='FavSectionBadge' Text='★ お気に入りフォルダ'
                                           FontSize='11' Foreground='#005FB8'
                                           VerticalAlignment='Center' Margin='6,0,0,0'
                                           Visibility='Collapsed'/>
                            </DockPanel>
                        </Border>

                        <Grid>
                            <ListView x:Name='ResultList' BorderThickness='0'
                                      Background='White' FontSize='12'
                                      SelectionMode='Extended'
                                      ItemContainerStyle='{StaticResource ResultItem}'>
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header='ファイル名' Width='240'>
                                            <GridViewColumn.CellTemplate><DataTemplate>
                                                <TextBlock Text='{Binding FileName}' ToolTip='{Binding FileName}'
                                                           TextTrimming='CharacterEllipsis' Margin='12,0,4,0'/>
                                            </DataTemplate></GridViewColumn.CellTemplate>
                                            <GridViewColumn.HeaderContainerStyle>
                                                <Style TargetType='GridViewColumnHeader' BasedOn='{StaticResource {x:Type GridViewColumnHeader}}'>
                                                    <Setter Property='Tag' Value='FileName'/>
                                                </Style>
                                            </GridViewColumn.HeaderContainerStyle>
                                        </GridViewColumn>
                                        <GridViewColumn Header='フォルダ' Width='180'>
                                            <GridViewColumn.CellTemplate><DataTemplate>
                                                <TextBlock Text='{Binding FolderPath}' ToolTip='{Binding FolderPath}'
                                                           TextTrimming='CharacterEllipsis' Margin='12,0,4,0'/>
                                            </DataTemplate></GridViewColumn.CellTemplate>
                                            <GridViewColumn.HeaderContainerStyle>
                                                <Style TargetType='GridViewColumnHeader' BasedOn='{StaticResource {x:Type GridViewColumnHeader}}'>
                                                    <Setter Property='Tag' Value='FolderPath'/>
                                                </Style>
                                            </GridViewColumn.HeaderContainerStyle>
                                        </GridViewColumn>
                                        <GridViewColumn Header='更新日 ▼' Width='130'>
                                            <GridViewColumn.CellTemplate><DataTemplate>
                                                <TextBlock Text='{Binding LastWrite}' HorizontalAlignment='Center'
                                                           FontFamily='Consolas'/>
                                            </DataTemplate></GridViewColumn.CellTemplate>
                                            <GridViewColumn.HeaderContainerStyle>
                                                <Style TargetType='GridViewColumnHeader' BasedOn='{StaticResource {x:Type GridViewColumnHeader}}'>
                                                    <Setter Property='Tag' Value='LastWriteTime'/>
                                                </Style>
                                            </GridViewColumn.HeaderContainerStyle>
                                        </GridViewColumn>
                                        <GridViewColumn Header='サイズ' Width='80'>
                                            <GridViewColumn.CellTemplate><DataTemplate>
                                                <TextBlock Text='{Binding Size}' HorizontalAlignment='Center'/>
                                            </DataTemplate></GridViewColumn.CellTemplate>
                                            <GridViewColumn.HeaderContainerStyle>
                                                <Style TargetType='GridViewColumnHeader' BasedOn='{StaticResource {x:Type GridViewColumnHeader}}'>
                                                    <Setter Property='Tag' Value='Size'/>
                                                </Style>
                                            </GridViewColumn.HeaderContainerStyle>
                                        </GridViewColumn>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                            <!-- ローディングオーバーレイ（データ行エリアのみ覆う） -->
                            <Border x:Name='SearchOverlay' Background='#CCFFFFFF'
                                    Visibility='Collapsed'>
                                <Path x:Name='SpinnerPath'
                                      Data='M 14,2 A 12,12 0 1 1 2,14'
                                      Stroke='#005FB8' StrokeThickness='2.5'
                                      StrokeStartLineCap='Round' StrokeEndLineCap='Round'
                                      Width='28' Height='28' Stretch='None'
                                      HorizontalAlignment='Center' VerticalAlignment='Center'
                                      RenderTransformOrigin='0.5,0.5'>
                                    <Path.RenderTransform><RotateTransform/></Path.RenderTransform>
                                </Path>
                            </Border>
                        </Grid>
                    </DockPanel>
                </Border>
            </DockPanel>

            <!-- === 右カラム（サイドパネル） === -->
            <Border Grid.Column='2' Background='White' BorderBrush='#E0E0E0'
                    BorderThickness='1' CornerRadius='4'>
                <DockPanel>
                    <!-- タブヘッダー（アンダーライン方式） -->
                    <Grid DockPanel.Dock='Top' Background='#FAFAFA'>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width='*'/><ColumnDefinition Width='*'/><ColumnDefinition Width='*'/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name='TabHistory' Grid.Column='0' Cursor='Hand'
                                Background='Transparent'
                                Padding='4,8,4,6' BorderThickness='0,0,0,2'
                                BorderBrush='#005FB8'>
                            <TextBlock x:Name='TabHistoryText' Text='検索履歴'
                                       FontSize='11' Foreground='#005FB8'
                                       FontWeight='Medium' HorizontalAlignment='Center'/>
                        </Border>
                        <Border x:Name='TabFile' Grid.Column='1' Cursor='Hand'
                                Background='Transparent'
                                Padding='4,8,4,6' BorderThickness='0,0,0,2'
                                BorderBrush='Transparent'>
                            <TextBlock x:Name='TabFileText' Text='ファイル'
                                       FontSize='11' Foreground='#777'
                                       HorizontalAlignment='Center'/>
                        </Border>
                        <Border x:Name='TabSeizure' Grid.Column='2' Cursor='Hand'
                                Background='Transparent'
                                Padding='4,8,4,6' BorderThickness='0,0,0,2'
                                BorderBrush='Transparent'>
                            <TextBlock x:Name='TabSeizureText' Text='差押登録'
                                       FontSize='11' Foreground='#777'
                                       HorizontalAlignment='Center'/>
                        </Border>
                    </Grid>
                    <!-- タブコンテンツ（Visibility で切替） -->
                    <Grid>
                        <ListBox x:Name='SearchHistoryList' BorderThickness='0'
                                 Background='Transparent'
                                 ItemContainerStyle='{StaticResource HistoryItem}'
                                 ScrollViewer.HorizontalScrollBarVisibility='Disabled'>
                            <ListBox.Template>
                                <ControlTemplate TargetType='ListBox'>
                                    <ScrollViewer Padding='0' Focusable='False'>
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </ControlTemplate>
                            </ListBox.Template>
                        </ListBox>
                        <ListBox x:Name='FileHistoryList' BorderThickness='0'
                                 Background='Transparent' Visibility='Collapsed'
                                 ItemContainerStyle='{StaticResource HistoryItem}'
                                 ScrollViewer.HorizontalScrollBarVisibility='Disabled'>
                            <ListBox.Template>
                                <ControlTemplate TargetType='ListBox'>
                                    <ScrollViewer Padding='0' Focusable='False'>
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </ControlTemplate>
                            </ListBox.Template>
                        </ListBox>
                        <ListBox x:Name='SeizureLogList' BorderThickness='0'
                                 Background='Transparent' Visibility='Collapsed'
                                 ItemContainerStyle='{StaticResource HistoryItem}'
                                 ScrollViewer.HorizontalScrollBarVisibility='Disabled'>
                            <ListBox.Template>
                                <ControlTemplate TargetType='ListBox'>
                                    <ScrollViewer Padding='0' Focusable='False'>
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </ControlTemplate>
                            </ListBox.Template>
                        </ListBox>
                    </Grid>
                </DockPanel>
            </Border>
        </Grid>
    </DockPanel>

    <!-- ============ お気に入りフォルダ設定オーバーレイ ============ -->
    <Grid x:Name='FavOverlay' Visibility='Collapsed'>
        <Border Background='#80000000'/>
        <Border x:Name='FavCard' Background='White' CornerRadius='8'
                Padding='20,20,24,20'
                HorizontalAlignment='Center' VerticalAlignment='Center'
                Width='440'>
            <StackPanel>
                <DockPanel Margin='0,0,0,4'>
                    <Button x:Name='FavCloseButton' DockPanel.Dock='Right'
                            Cursor='Hand' Focusable='False'
                            VerticalAlignment='Top' Margin='0,-4,-4,0'>
                        <Button.Template>
                            <ControlTemplate TargetType='Button'>
                                <Border x:Name='cbg' Width='24' Height='24'
                                        CornerRadius='4' Background='Transparent'>
                                    <TextBlock Text='✕' FontSize='13'
                                               Foreground='#999'
                                               HorizontalAlignment='Center'
                                               VerticalAlignment='Center'/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property='IsMouseOver' Value='True'>
                                        <Setter TargetName='cbg' Property='Background' Value='#F0F0F0'/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <TextBlock Text='お気に入りフォルダの設定'
                               FontSize='15' FontWeight='Medium'/>
                </DockPanel>
                <TextBlock Text='最大 5 フォルダまで登録できます。サブフォルダも再帰的に検索します。'
                           FontSize='11' Foreground='#999' Margin='0,0,0,16'/>
                <StackPanel x:Name='FavSlotPanel'/>
                <Border BorderBrush='#E0E0E0' BorderThickness='0,1,0,0'
                        Padding='0,12,0,0' Margin='0,6,0,0'>
                    <StackPanel>
                        <TextBlock Text='フォルダを追加' FontSize='11'
                                   Foreground='#999' Margin='0,0,0,6'/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width='*'/>
                                <ColumnDefinition Width='6'/>
                                <ColumnDefinition Width='Auto'/>
                                <ColumnDefinition Width='6'/>
                                <ColumnDefinition Width='Auto'/>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column='0' Background='White'
                                    BorderBrush='#D0D0D0' BorderThickness='1'
                                    CornerRadius='4'>
                                <TextBox x:Name='FavPathInput' FontSize='11'
                                         FontFamily='Consolas'
                                         Padding='8,6' BorderThickness='0'
                                         Background='Transparent'
                                         VerticalContentAlignment='Center'/>
                            </Border>
                            <Button x:Name='FavBrowseButton' Grid.Column='2'
                                    Style='{StaticResource GB}'
                                    Padding='8,6' FontSize='11'>
                                <StackPanel Orientation='Horizontal'>
                                    <Canvas Width='14' Height='14' Margin='0,0,4,0'>
                                        <Path Data='M1,4 L1,12 L13,12 L13,6 L6.5,6 L5,4 Z'
                                              Stroke='#555' StrokeThickness='1' Fill='#E8EEF4'/>
                                    </Canvas>
                                    <TextBlock Text='参照' VerticalAlignment='Center'/>
                                </StackPanel>
                            </Button>
                            <Button x:Name='FavAddButton' Grid.Column='4'
                                    Style='{StaticResource AB}'
                                    Padding='8,6' FontSize='11'>
                                <StackPanel Orientation='Horizontal'>
                                    <TextBlock Text='＋' FontWeight='Bold'
                                               Margin='0,0,3,0'/>
                                    <TextBlock Text='追加'/>
                                </StackPanel>
                            </Button>
                        </Grid>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>
    </Grid>

    <!-- ============ アップデート告知オーバーレイ ============ -->
    <Grid x:Name='NoticeOverlay' Visibility='Collapsed'>
        <Border Background='#80000000'/>
        <Border x:Name='NoticeCard' Background='White' CornerRadius='8'
                Padding='24,24,28,20'
                HorizontalAlignment='Center' VerticalAlignment='Center'
                Width='440'>
            <StackPanel>
                <DockPanel Margin='0,0,0,0'>
                    <Button x:Name='NoticeCloseButton' DockPanel.Dock='Right'
                            Cursor='Hand' Focusable='False'
                            VerticalAlignment='Top' Margin='0,-4,-4,0'>
                        <Button.Template>
                            <ControlTemplate TargetType='Button'>
                                <Border x:Name='ncbg' Width='24' Height='24'
                                        CornerRadius='4' Background='Transparent'>
                                    <TextBlock Text='✕' FontSize='13'
                                               Foreground='#999'
                                               HorizontalAlignment='Center'
                                               VerticalAlignment='Center'/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property='IsMouseOver' Value='True'>
                                        <Setter TargetName='ncbg' Property='Background' Value='#F0F0F0'/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Button.Template>
                    </Button>
                    <TextBlock Text='アップデートのお知らせ'
                               FontSize='15' FontWeight='Medium'/>
                </DockPanel>
                <StackPanel x:Name='NoticeContent' Margin='0,4,0,0'/>
            </StackPanel>
        </Border>
    </Grid>

    </Grid>
</Window>";

        using (var reader = XmlReader.Create(new StringReader(xaml)))
        {
            return (Window)XamlReader.Load(reader);
        }
    }
}

// ==============================================================
// RelayCommand（キーボードショートカット用の ICommand 実装）
// ==============================================================

public class RelayCommand : ICommand
{
    private Action<object> execute;
    public RelayCommand(Action<object> execute) { this.execute = execute; }
    public event EventHandler CanExecuteChanged { add { } remove { } }
    public bool CanExecute(object parameter) { return true; }
    public void Execute(object parameter) { execute(parameter); }
}