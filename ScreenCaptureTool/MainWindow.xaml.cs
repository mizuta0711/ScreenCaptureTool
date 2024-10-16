using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.VisualBasic.FileIO;

using PixelFormat = System.Drawing.Imaging.PixelFormat;
using MessageBox = System.Windows.MessageBox;
using System.Text;

namespace ScreenCaptureTool
{
    public partial class MainWindow : Window
    {
        #region Win32 APIのインポート

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        private static extern int BitBlt(IntPtr hdcDest, int xDest, int yDest, int wDest, int hDest, IntPtr hdcSrc, int xSrc, int ySrc, int rop);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        #endregion Win32 APIのインポート

        #region Variable

        /// <summary>
        /// ファイル名の配列
        /// </summary>
        private ObservableCollection<string> SaveFileNames = new ObservableCollection<string>();

        /// <summary>
        /// 画像ファイルのリスト
        /// </summary>
        public ObservableCollection<ImageFile> ImageFiles { get; set; } = new ObservableCollection<ImageFile>();

        /// <summary>
        /// サムネイルサイズ
        /// </summary>
        private int thumbnailSize = 200;

        /// <summary>
        /// 保存先パス
        /// </summary>
        private string saveFolderPath = Path.Combine(Environment.CurrentDirectory, "CapturedImages");

        /// <summary>
        /// 現在のプロジェクト設定
        /// </summary>
        private ProjectSettings projectSettings = new ProjectSettings();

        #endregion Variable

        #region Constructor

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            // DataContextにImageFilesをバインド
            DataContext = this;

            // ラジオボタンの選択によって表示するUIを切り替える
            CaptureRectRadioButton.Checked += CaptureOption_CheckedChanged;
            CaptureWindowRadioButton.Checked += CaptureOption_CheckedChanged;

            // フォルダツリーの初期化
            InitializeFolderTree();

            // プロジェクトファイルを読み込む
            LoadProjectFile(projectSettings.FilePath);
        }

        #endregion Constructor

        #region Methods

        #region Events(Window)

        /// <summary>
        /// ウィンドウが閉じられた：ウィンドウの位置とサイズを保存する
        /// </summary>
        /// <param name="e"></param>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // プロジェクトファイルを保存
            SaveProjectFile(projectSettings);
        }

        #endregion Events(Window)

        #region Private

        #region Settings

        /// <summary>
        /// プロジェクト設定をUIに反映
        /// </summary>
        /// <param name="settings">プロジェクト設定</param>
        private void LoadProjectSettings(ProjectSettings settings)
        {
            // キャプチャー範囲
            CaptureLeftTextBox.Text = settings.CaptureLeft.ToString();
            CaptureTopTextBox.Text = settings.CaptureTop.ToString();
            CaptureWidthTextBox.Text = settings.CaptureWidth.ToString();
            CaptureHeightTextBox.Text = settings.CaptureHeight.ToString();

            // サムネイルサイズ
            if (settings.ThumbnailSize > 0)
            {
                thumbnailSize = settings.ThumbnailSize;
            }

            // ウィンドウの位置とサイズを設定
            if (settings.WindowTop >= 0 && settings.WindowLeft >= 0)
            {
                Top = settings.WindowTop;
                Left = settings.WindowLeft;
            }
            if (settings.WindowWidth > 0 && settings.WindowHeight > 0)
            {
                Width = settings.WindowWidth;
                Height = settings.WindowHeight;
            }

            // ウィンドウタイトルを復元
            WindowTitleTextBox.Text = settings.CaptureWindowTitle;

            // チャプチャータイプを復元
            if (settings.SelectedCaptureType == ProjectSettings.CaptureType.ScreenRect)
            {
                CaptureRectRadioButton.IsChecked = true;
            }
            else
            {
                CaptureWindowRadioButton.IsChecked = true;
            }

            // 保存ファイル名一覧
            SaveFileNames = settings.SaveFileNames;

            // 保存先フォルダ
            saveFolderPath = settings.SaveFolderPath;

            // UIに反映
            // 保存先フォルダ名
            SelectedFolderTextBox.Text = saveFolderPath;

            // ファイル一覧
            FileNameComboBox.ItemsSource = SaveFileNames;

            // 保存先フォルダをツリーから選択状態にする
            SelectFolderInTree(saveFolderPath);
        }

        /// <summary>
        /// UIの設定をプロジェクト設定に保存する
        /// </summary>
        /// <param name="settings">プロジェクト設定</param>
        private void StoreProjectSettings(ProjectSettings settings)
        {
            // ウィンドウの位置とサイズ
            settings.WindowTop = Top;
            settings.WindowLeft = Left;
            settings.WindowWidth = Width;
            settings.WindowHeight = Height;
            // キャプチャー範囲
            settings.CaptureLeft = int.TryParse(CaptureLeftTextBox.Text, out var x) ? x : 0;
            settings.CaptureTop = int.TryParse(CaptureTopTextBox.Text, out var y) ? y : 0;
            settings.CaptureWidth = int.TryParse(CaptureWidthTextBox.Text, out var width) ? width : 0;
            settings.CaptureHeight = int.TryParse(CaptureHeightTextBox.Text, out var height) ? height : 0;
            // キャプチャーウィンドウタイトル
            settings.CaptureWindowTitle = WindowTitleTextBox.Text;
            // チャプチャータイプ
            settings.SelectedCaptureType = (CaptureRectRadioButton.IsChecked ?? true) ? ProjectSettings.CaptureType.ScreenRect : ProjectSettings.CaptureType.Window;
            // サムネイルサイズ
            settings.ThumbnailSize = thumbnailSize;
            // 保存ファイル名一覧
            settings.SaveFileNames = SaveFileNames;
            // 保存先フォルダ
            settings.SaveFolderPath = saveFolderPath;
        }

        /// <summary>
        /// プロジェクト設定をファイルから読み込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>true: 成功 / false: 失敗</returns>
        private bool LoadProjectFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                ShowErrorDialog("指定されたプロジェクトファイルが見つかりません: " + filePath);
                return false;
            }

            var settings = ProjectSettings.Load(filePath);
            if (settings == null)
            {
                ShowErrorDialog("設定ファイルの読み込みに失敗しました");
                return false;
            }

            // プロジェクト設定をUIに反映
            projectSettings = settings;
            LoadProjectSettings(projectSettings);
            return true;
        }

        /// <summary>
        /// プロジェクト設定をファイルに保存する
        /// </summary>
        /// <param name="settings">プロジェクト設定</param>
        /// <returns>true: 成功 / false: 失敗</returns>
        private bool SaveProjectFile(ProjectSettings settings)
        {
            // UIの設定をプロジェクト設定に反映
            StoreProjectSettings(settings);

            // ファイルパスが未設定の場合はカレントディレクトリに保存
            if (settings.FilePath == null)
            {
                settings.FilePath = Path.Combine(Environment.CurrentDirectory, "ScreenCaptureTool.ssp");
            }

            // ファイルに保存
            if (settings.Save(settings.FilePath) == false)
            {
                ShowErrorDialog("設定ファイルの保存に失敗しました");
                return false;
            }

            return true;
        }

        private void LoadProjectButton_Click(object sender, RoutedEventArgs e)
        {
            // ダイアログでファイルを選択
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Screen Capture Project (*.scp)|*.scp";
            if (openFileDialog.ShowDialog() == true)
            {
                LoadProjectFile(openFileDialog.FileName);
            }
        }

        private void SaveAsProjectButton_Click(object sender, RoutedEventArgs e)
        {
            // ダイアログで保存先を指定
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Screen Capture Project (*.scp)|*.scp";
            if (saveFileDialog.ShowDialog() == true)
            {
                projectSettings.FilePath = saveFileDialog.FileName;
                SaveProjectFile(projectSettings);
            }
        }

        #endregion Settings

        #region Utility

        /// <summary>
        /// ファイルがロックされているか
        /// </summary>
        /// <param name="filePath">パス</param>
        /// <returns>true:ロックされている　/ false:ロックされていない</returns>
        private static bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    // ファイルはロックされていない
                }
            }
            catch (IOException)
            {
                // ファイルはロックされている
                return true;
            }
            return false;
        }

        /// <summary>
        /// VisualTreeHelperを使って特定の型の子要素を取得する汎用メソッド
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                {
                    return typedChild;
                }
                else
                {
                    T? childOfChild = FindVisualChild<T>(child);
                    if (childOfChild != null)
                    {
                        return childOfChild;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// エラーダイアログを表示する
        /// </summary>
        /// <param name="message">メッセージ</param>
        /// <param name="title">タイトル</param>
        private static void ShowErrorDialog(string message)
        {
            MessageBox.Show(message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /// <summary>
        /// 情報ダイアログを表示する
        /// </summary>
        /// <param name="message">メッセージ</param>
        /// <param name="title">タイトル</param>
        private static void ShowInformationDialog(string message)
        {
            MessageBox.Show(message, "情報", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion Utility

        #region BitmapUtility

        /// <summary>
        /// BitmapSourceからBitmapImageへの変換
        /// </summary>
        /// <param name="bitmapSource">BitmapSource</param>
        /// <returns>変換後のBitmapImage</returns>
        private static BitmapImage ConvertBitmapSourceToBitmapImage(BitmapSource bitmapSource)
        {
            // BitmapSourceをMemoryStreamに保存
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // BitmapEncoderでBitmapSourceをエンコード
                BitmapEncoder encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                encoder.Save(memoryStream);

                // MemoryStreamからBitmapImageを作成
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = new MemoryStream(memoryStream.ToArray());
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        /// <summary>
        /// 指定されたファイルをBitmapImage形式で読み込む
        /// </summary>
        /// <param name="filePath">パス</param>
        /// <returns>BitmapImage</returns>
        private static BitmapImage LoadBitmapImage(string filePath)
        {
            using (Stream stream = new FileStream(
                filePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete
            ))
            {
                // ロックしないように指定したstreamを使用する。
                BitmapDecoder decoder = BitmapDecoder.Create(
                    stream,
                    BitmapCreateOptions.None, // この辺のオプションは適宜
                    BitmapCacheOption.Default // これも
                );
                BitmapSource bmp = new WriteableBitmap(decoder.Frames[0]);
                bmp.Freeze();

                // BitmapImage形式に変換して返す
                return ConvertBitmapSourceToBitmapImage(bmp);
            }
        }

        #endregion BitmapUtility

        #region CaptureTools

        /// <summary>
        /// ファイルのフルパスを作成
        /// </summary>
        /// <param name="folderPath">フォルダ名</param>
        /// <param name="fileName">ファイル名(拡張子を除く)</param>
        /// <param name="extension">拡張子(デフォルト："png")</param>
        /// <returns>フルパス</returns>
        private string CreateFilePath(string folderPath, string fileName, string extension = "png")
        {
            // パスを作成
            string fullFileName = $"{fileName}.{extension}";
            return Path.Combine(folderPath, fullFileName);
        }

        /// <summary>
        /// BitmapをPNG形式で保存する
        /// </summary>
        /// <param name="bitmap">Bitmap</param>
        /// <param name="filePath">保存先のパス</param>
        /// <returns>true: 保存 / false: 失敗</returns>
        private bool SaveBitmapAsPng(Bitmap bitmap, string filePath)
        {
            // フォルダが存在しない場合は作成する
            if (!Directory.Exists(Path.GetDirectoryName(filePath)))
            {
                if (Path.GetDirectoryName(filePath) is string parentFolderPath)
                {
                    Directory.CreateDirectory(parentFolderPath);
                }
            }

            // ファイルが既に存在する場合、上書き確認ダイアログを表示
            if (File.Exists(filePath))
            {
                var result = MessageBox.Show(
                    "このファイルは既に存在します。上書きしますか？",
                    "上書き確認",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    // 上書きをキャンセル
                    return false;
                }

                // 既存ファイルを削除
                if (DeleteImageFile(filePath) == false)
                {
                    // 上書きをキャンセル
                    return false;
                }
            }

            // PNGとして保存
            try
            {
                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                return true;
            }
            catch (Exception ex)
            {
                ShowErrorDialog("ファイルの保存に失敗しました: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// 画像ファイルを削除(ゴミ箱に移動)
        /// ※サムネイル一覧も更新する
        /// </summary>
        /// <param name="filePath">パス</param>
        /// <param name="isMoveToTrush">ゴミ箱に移動するか</param>
        /// <returns>true:成功 / false:失敗</returns>
        private bool DeleteImageFile(string filePath, bool isMoveToTrush = true)
        {
            // 実際のファイルも削除
            if (File.Exists(filePath))
            {
                // ファイルがロックされているかどうかを確認
                if (!IsFileLocked(filePath))
                {
                    try
                    {
                        if (isMoveToTrush)
                        {
                            // ファイルをゴミ箱に移動
                            FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                        }
                        else
                        {
                            // ファイルを削除
                            File.Delete(filePath);
                        }
                    }
                    catch (Exception ex)
                    {
                        ShowErrorDialog("ファイルの削除に失敗しました: " + ex.Message);
                        return false;
                    }
                }
                else
                {
                    ShowErrorDialog("ファイルがロックされているため、削除できません。");
                    return false;
                }
            }

            // 保存先フォルダと削除対象のファイルが同一フォルダか？
            if (Path.GetDirectoryName(filePath) == saveFolderPath)
            {
                // 一致する ImageFile を検索
                var itemToRemove = ImageFiles.FirstOrDefault(item => item.FileName == Path.GetFileName(filePath));
                if (itemToRemove != null)
                {
                    // コレクションから削除
                    ImageFiles.Remove(itemToRemove);
                }
            }

            return true;
        }

        /// <summary>
        /// キャプチャー画像を現在選択されているファイル名で保存する
        /// ※サムネイル一覧も更新する
        /// </summary>
        /// <param name="bitmap">画像</param>
        /// <returns>true: 成功 / false: 失敗</returns>
        private bool SaveCaptureImage(Bitmap bitmap)
        {
            // ComboBoxから選択または入力されたファイル名を取得
            string selectedFileName = FileNameComboBox.Text.Trim();
            if (string.IsNullOrEmpty(selectedFileName))
            {
                ShowErrorDialog("ファイル名を入力してください。");
                return false;
            }

            // PNGで保存
            string filePath = CreateFilePath(saveFolderPath, selectedFileName);
            if (SaveBitmapAsPng(bitmap, filePath) == false)
            {
                return false;
            }

            // サムネイルリストを更新
            AddImageToList(filePath);

            // 新しいファイル名をComboBoxのリストに追加
            if (!FileNameComboBox.Items.Contains(selectedFileName))
            {
                SaveFileNames.Add(selectedFileName);
            }

            return true;
        }

        #endregion CaptureTools

        #region ScreenCapture

        /// <summary>
        /// デスクトップの指定範囲をキャプチャーする
        /// </summary>
        /// <returns>true: 成功　/ false: 失敗</returns>
        private bool CaptureScreenRect()
        {
            // デスクトップの解像度を取得
            int screenWidth = (int)SystemParameters.VirtualScreenWidth;
            int screenHeight = (int)SystemParameters.VirtualScreenHeight;

            // UIから矩形の設定を取得
            if (!int.TryParse(CaptureLeftTextBox.Text, out int x) ||
                !int.TryParse(CaptureTopTextBox.Text, out int y) ||
                !int.TryParse(CaptureWidthTextBox.Text, out int width) ||
                !int.TryParse(CaptureHeightTextBox.Text, out int height))
            {
                ShowErrorDialog("矩形の設定が無効です。X、Y、幅、高さを正しく入力してください。");
                return false;
            }

            // 矩形のサイズをチェックして調整
            if (x < 0 || y < 0 || width <= 0 || height <= 0 ||
                x + width > screenWidth || y + height > screenHeight)
            {
                ShowErrorDialog("指定された矩形が無効です。画面の範囲内で正しい矩形を指定してください。");
                return false;
            }

            // 矩形のビットマップを作成
            using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                }
                return SaveCaptureImage(bitmap);
            }
        }

        #endregion ScreenCapture

        #region CaptureWindow

        // 部分一致でウィンドウを検索
        private IntPtr FindWindowByTitle(string partialTitle)
        {
            IntPtr foundWindow = IntPtr.Zero;

            EnumWindows((hWnd, lParam) =>
            {
                StringBuilder windowTitle = new StringBuilder(256);
                GetWindowText(hWnd, windowTitle, 256);

                if (windowTitle.ToString().Contains(partialTitle, StringComparison.OrdinalIgnoreCase) && IsWindowVisible(hWnd))
                {
                    foundWindow = hWnd;
                    return false; // ウィンドウが見つかったので列挙を終了
                }
                return true; // まだ見つかっていないので続行
            }, IntPtr.Zero);

            return foundWindow;
        }

        /// <summary>
        /// ウィンドウタイトルキャプチャ処理
        /// </summary>
        /// <param name="windowTitle">ウィンドウタイトル</param>
        /// <returns>true: 成功　/ false: 失敗</returns>
        private bool CaptureWindowByTitle(string windowTitle)
        {
            if (string.IsNullOrWhiteSpace(windowTitle))
            {
                ShowErrorDialog("ウィンドウタイトルを入力してください。");
                return false;
            }

            IntPtr hWnd = FindWindowByTitle(windowTitle);
            if (hWnd == IntPtr.Zero)
            {
                ShowErrorDialog("指定されたウィンドウが見つかりません。");
                return false;
            }

            // ウィンドウの位置とサイズを取得
            if (GetWindowRect(hWnd, out RECT rect) == false)
            {
                ShowErrorDialog("ウィンドウの位置を取得できませんでした。");
                return false;
            }

            // RECTから幅と高さを計算
            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;

            // ウィンドウ全体のビットマップを作成
            Bitmap bitmap = CaptureWindow(hWnd, width, height);
            return SaveCaptureImage(bitmap);
        }

        /// <summary>
        /// ウィンドウをキャプチャーしたBitmapを生成
        /// </summary>
        /// <param name="hWnd">ウィンドウハンドル</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        /// <returns>Bitmap</returns>
        private Bitmap CaptureWindow(IntPtr hWnd, int width, int height)
        {
            const int SRCCOPY = 0x00CC0020;

            // ウィンドウのDCを取得
            IntPtr hdcWindow = GetDC(hWnd);
            IntPtr hdcMemDC = CreateCompatibleDC(hdcWindow);

            // ウィンドウのビットマップを作成
            IntPtr hBitmap = CreateCompatibleBitmap(hdcWindow, width, height);
            IntPtr hOld = SelectObject(hdcMemDC, hBitmap);

            // ウィンドウのビットブロック転送 (BitBlt) を実行
            BitBlt(hdcMemDC, 0, 0, width, height, hdcWindow, 0, 0, SRCCOPY);

            // ビットマップを取得
            Bitmap bmp = Bitmap.FromHbitmap(hBitmap);

            // リソース解放
            SelectObject(hdcMemDC, hOld);
            DeleteObject(hBitmap);
            DeleteDC(hdcMemDC);

            // ウィンドウのDCを解放
            DeleteDC(hdcWindow);

            return bmp;
        }

        #endregion CaptureWindow

        #region Thumbnails

        /// <summary>
        /// 保存先フォルダの画像一覧を取得する
        /// </summary>
        private void LoadImagesFromFolder()
        {
            ImageFiles.Clear();

            if (Directory.Exists(saveFolderPath))
            {
                var files = Directory.GetFiles(saveFolderPath, "*.png");
                foreach (var file in files)
                {
                    AddImageToList(file);
                }
            }
        }

        /// <summary>
        /// 画像を読み込んで画像一覧に追加する
        /// </summary>
        /// <param name="filePath">画像のパス</param>
        private void AddImageToList(string filePath)
        {
            // 保存先フォルダと削除対象のファイルが同一フォルダでない場合は追加しない
            if (Path.GetDirectoryName(filePath) != saveFolderPath)
            {
                return;
            }

            try
            {
                // サムネル画像を読み込む
                BitmapImage thumbnail = LoadBitmapImage(filePath);

                // xamlでImageを記述 → imgSync
                // 画像ファイル情報をリストに追加
                ImageFiles.Add(new ImageFile
                {
                    FileName = Path.GetFileName(filePath),
                    Thumbnail = thumbnail,
                    ThumbnailWidth = thumbnailSize,
                    ThumbnailHeight = thumbnailSize
                });
            }
            catch (Exception ex)
            {
                ShowErrorDialog("画像の読み込みに失敗しました: " + ex.Message);
            }
        }

        /// <summary>
        /// サムネイル画像のサイズを更新する
        /// </summary>
        private void UpdateThumbnailsSize()
        {
            foreach (var imageFile in ImageFiles)
            {
                // サムネイルの幅と高さを新しいサイズに合わせて更新
                imageFile.ThumbnailWidth = thumbnailSize;
                imageFile.ThumbnailHeight = thumbnailSize;
            }
        }

        #endregion Thumbnails

        #region FolderTree

        // フォルダツリーを初期化する
        private void InitializeFolderTree()
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady)
                {
                    var item = new TreeViewItem { Header = drive.Name, Tag = drive.Name };
                    item.Items.Add(null);  // ダミーアイテム
                    item.Expanded += Folder_Expanded;
                    FolderTreeView.Items.Add(item);
                }
            }
        }

        // フォルダを展開したときの処理
        private void Folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == null)  // ダミーアイテムの確認
            {
                item.Items.Clear();
                try
                {
                    var directories = Directory.GetDirectories(item.Tag.ToString());
                    foreach (var directory in directories)
                    {
                        var subItem = new TreeViewItem { Header = Path.GetFileName(directory), Tag = directory };
                        subItem.Items.Add(null);  // ダミーアイテム
                        subItem.Expanded += Folder_Expanded;
                        item.Items.Add(subItem);
                    }
                }
                catch (UnauthorizedAccessException) { }
            }
        }

        // フォルダツリーで選択が変更されたとき
        private void FolderTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var selectedItem = FolderTreeView.SelectedItem as TreeViewItem;
            if (selectedItem != null)
            {
                if (selectedItem.Tag.ToString() is string folderPath)
                {
                    SelectCurrentFolder(folderPath);
                }
            }
        }

        // saveFolderPathのフォルダをツリーで選択する
        private void SelectFolderInTree(string saveFolderPath)
        {
            string[] pathParts = saveFolderPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            TreeViewItem? currentItem = null;

            foreach (TreeViewItem driveItem in FolderTreeView.Items)
            {
                if (driveItem.Tag.ToString() == pathParts[0] + Path.DirectorySeparatorChar) // ドライブ名の一致を確認
                {
                    currentItem = driveItem;
                    currentItem.IsExpanded = true; // ドライブを展開
                    break;
                }
            }

            // ドライブが見つからない場合は終了
            if (currentItem == null)
            {
                return;
            }

            // ドライブ以下のフォルダを順次展開していく
            for (int i = 1; i < pathParts.Length; i++)
            {
                bool found = false;
                foreach (TreeViewItem subItem in currentItem.Items)
                {
                    if (subItem.Header.ToString() == pathParts[i])
                    {
                        currentItem = subItem;
                        currentItem.IsExpanded = true; // フォルダを展開
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    return; // 見つからなければ終了
                }
            }

            currentItem.IsSelected = true; // 最後のフォルダを選択
        }

        /// <summary>
        /// カレントフォルダの変更
        /// </summary>
        /// <param name="folderPath">パス</param>
        private void SelectCurrentFolder(string folderPath)
        {
            // 保存先フォルダを変更
            saveFolderPath = folderPath;

            // UIに表示
            SelectedFolderTextBox.Text = folderPath;

            // 選択されたフォルダの一覧を表示
            LoadImagesFromFolder();

            // 選択されたフォルダのツリー更新
            RefreshSelectedFolderTree();
        }

        /// <summary>
        /// フォルダを再探索してツリーを更新する
        /// </summary>
        private void RefreshSelectedFolderTree()
        {
            var selectedItem = FolderTreeView.SelectedItem as TreeViewItem;
            if (selectedItem != null)
            {
                // 現在選択されたフォルダの子要素をクリア
                selectedItem.Items.Clear();

                // 再度フォルダを展開し、ツリーに反映
                try
                {
                    if (selectedItem.Tag.ToString() is string folderPath)
                    {
                        var directories = Directory.GetDirectories(folderPath);
                        foreach (var directory in directories)
                        {
                            var subItem = new TreeViewItem { Header = Path.GetFileName(directory), Tag = directory };
                            subItem.Items.Add(null);  // ダミーアイテム
                            subItem.Expanded += Folder_Expanded;
                            selectedItem.Items.Add(subItem);
                        }
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    // アクセスできないフォルダに対してのエラーハンドリング
                    ShowErrorDialog("フォルダにアクセスできませんでした: " + ex.Message);
                }
            }
        }

        #endregion FolderTree

        #endregion Private

        #region Events(Control)

        /// <summary>
        /// フォルダ選択ボタン：押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SelectCurrentFolder(dialog.SelectedPath);
                }
            }
        }

        /// <summary>
        /// キャプチャー選択ラジオボタン：選択が変更された
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CaptureOption_CheckedChanged(object sender, RoutedEventArgs e)
        {
            // 矩形キャプチャーが選ばれている場合、矩形入力パネルを表示
            if (CaptureRectRadioButton.IsChecked == true)
            {
                RectCapturePanel.Visibility = Visibility.Visible;
                WindowCapturePanel.Visibility = Visibility.Collapsed;
            }
            else if (CaptureWindowRadioButton.IsChecked == true)
            {
                RectCapturePanel.Visibility = Visibility.Collapsed;
                WindowCapturePanel.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// サムネイルサイズ変更のコンボボックス：選択変更
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThumbnailSizeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ThumbnailSizeComboBox.SelectedItem is ComboBoxItem selectedItem)
            {
                // ComboBoxのTagからサイズを取得
                thumbnailSize = int.Parse((string)selectedItem.Tag);
                UpdateThumbnailsSize();  // サイズ変更後にサムネイル更新
            }
        }

        /// <summary>
        /// 撮影ボタン：押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CaptureButton_Click(object sender, RoutedEventArgs e)
        {
            // ラジオボタンで選択されたキャプチャ方法に応じて処理を分ける
            if (CaptureRectRadioButton.IsChecked == true)
            {
                // 矩形キャプチャ処理
                CaptureScreenRect();
            }
            else if (CaptureWindowRadioButton.IsChecked == true)
            {
                // ウィンドウタイトルキャプチャ処理
                CaptureWindowByTitle(WindowTitleTextBox.Text);
            }
        }

        /// <summary>
        /// サムネイル画像：左クリック
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Thumbnail_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // ダブルクリックかどうかを確認
            if (e.ClickCount == 2)
            {
                // サムネイルのイメージを取得
                var image = sender as System.Windows.Controls.Image;
                if (image != null)
                {
                    // バインドされた ImageFile オブジェクトを取得
                    var selectedImageFile = image.DataContext as ImageFile;
                    if (selectedImageFile != null)
                    {
                        // フルパスをProcess.Startに渡す
                        string fullPath = System.IO.Path.Combine(saveFolderPath, selectedImageFile.FileName);

                        // フォトアプリで画像を開く
                        Process.Start(new ProcessStartInfo(fullPath)
                        {
                            UseShellExecute = true // 既定のアプリケーションで開く
                        });
                    }
                }
            }
        }

        /// <summary>
        /// 削除メニュー：選択
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 選択されたサムネイルを取得
            var menuItem = sender as MenuItem;
            var imageFile = menuItem.DataContext as ImageFile;

            if (imageFile == null) return;

            string filePath = Path.Combine(saveFolderPath, imageFile.FileName);

            if (IsFileLocked(filePath))
            {
                ShowErrorDialog("ファイルがロックされているため、削除できません。");
                return;
            }

            // 確認ダイアログを表示
            if (MessageBox.Show($"このファイルを削除しますか？\n{filePath}", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                try
                {
                    // ファイルをゴミ箱に移動
                    FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);

                    // リストから削除
                    ImageFiles.Remove(imageFile);

                    ShowInformationDialog("ファイルはゴミ箱に移動されました。");
                }
                catch (Exception ex)
                {
                    ShowErrorDialog("ファイルの削除に失敗しました: " + ex.Message);
                }
            }
        }

        #endregion Events(Control)

        #endregion Methods
    }
}