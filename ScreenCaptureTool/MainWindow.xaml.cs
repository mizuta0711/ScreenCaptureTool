using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using System.Xml.Serialization;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Diagnostics;

using PixelFormat = System.Drawing.Imaging.PixelFormat;
using MessageBox = System.Windows.MessageBox;
using Microsoft.VisualBasic.FileIO;
using System.Linq;

namespace ScreenCaptureTool
{
    public partial class MainWindow : Window
    {
        // 画像ファイルのリスト
        public ObservableCollection<ImageFile> ImageFiles { get; set; } = new ObservableCollection<ImageFile>();

        // デフォルトのファイル名の配列
        private readonly List<string> defaultFileNames = new List<string>() {
            "フロー", "期待値1", "期待値2", "期待値3", "条件1", "条件2", "条件3"
        };

        private string saveFolderPath = System.IO.Path.Combine(Environment.CurrentDirectory, "CapturedImages");

        private const string SettingsFilePath = "settings.xml";

        public MainWindow()
        {
            InitializeComponent();

            // DataContextにImageFilesをバインド
            DataContext = this;

            // ComboBox にファイル名リストを設定
            FileNameComboBox.ItemsSource = defaultFileNames;

            // 設定を読み込む
            LoadSettings();

            // 初期の保存先フォルダを表示
            SelectedFolderPath.Text = "選択されたフォルダ: " + saveFolderPath;

            // 起動時に画像フォルダをチェック
            LoadImagesFromFolder();
        }

        private void LoadSettings()
        {
            if (File.Exists(SettingsFilePath))
            {
                try
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(AppSettings));
                    using (FileStream fs = new FileStream(SettingsFilePath, FileMode.Open))
                    {
                        AppSettings settings = (AppSettings)serializer.Deserialize(fs);
                        saveFolderPath = settings.SaveFolderPath;
                        XInput.Text = settings.X.ToString();
                        YInput.Text = settings.Y.ToString();
                        WidthInput.Text = settings.Width.ToString();
                        HeightInput.Text = settings.Height.ToString();
                        SelectedFolderPath.Text = "選択されたフォルダ: " + saveFolderPath;

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
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("設定ファイルの読み込みに失敗しました: " + ex.Message);
                }
            }
            else
            {
                SelectedFolderPath.Text = "選択されたフォルダ: " + saveFolderPath;
            }
        }

        private void SaveSettings()
        {
            try
            {
                AppSettings settings = new AppSettings
                {
                    SaveFolderPath = saveFolderPath,
                    X = int.TryParse(XInput.Text, out var x) ? x : 0,
                    Y = int.TryParse(YInput.Text, out var y) ? y : 0,
                    Width = int.TryParse(WidthInput.Text, out var width) ? width : 0,
                    Height = int.TryParse(HeightInput.Text, out var height) ? height : 0,
                    WindowTop = Top,
                    WindowLeft = Left,
                    WindowWidth = Width,
                    WindowHeight = Height
                };

                XmlSerializer serializer = new XmlSerializer(typeof(AppSettings));
                using (FileStream fs = new FileStream(SettingsFilePath, FileMode.Create))
                {
                    serializer.Serialize(fs, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("設定ファイルの保存に失敗しました: " + ex.Message);
            }
        }

        private void CaptureDesktopButton_Click(object sender, RoutedEventArgs e)
        {
            // デスクトップ全体をキャプチャ
            CaptureDesktop();
        }

        private void CaptureDesktop()
        {
            // デスクトップの解像度を取得
            int screenWidth = (int)SystemParameters.VirtualScreenWidth;
            int screenHeight = (int)SystemParameters.VirtualScreenHeight;

            // UIから矩形の設定を取得
            if (!int.TryParse(XInput.Text, out int x) ||
                !int.TryParse(YInput.Text, out int y) ||
                !int.TryParse(WidthInput.Text, out int width) ||
                !int.TryParse(HeightInput.Text, out int height))
            {
                MessageBox.Show("矩形の設定が無効です。X、Y、幅、高さを正しく入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 矩形のサイズをチェックして調整
            if (x < 0 || y < 0 || width <= 0 || height <= 0 ||
                x + width > screenWidth || y + height > screenHeight)
            {
                MessageBox.Show("指定された矩形が無効です。画面の範囲内で正しい矩形を指定してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 矩形のビットマップを作成
            using (Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                }

                // ComboBoxから選択または入力されたファイル名を取得
                string selectedFileName = FileNameComboBox.Text.Trim();
                if (string.IsNullOrEmpty(selectedFileName))
                {
                    MessageBox.Show("ファイル名を入力してください。");
                    return;
                }
                string filePath = SaveBitmapAsPng(bitmap, selectedFileName);

                if (filePath != null)
                {
                    // サムネイルリストを更新
                    AddImageToList(filePath);

                    // 新しいファイル名をComboBoxのリストに追加
                    if (!FileNameComboBox.Items.Contains(selectedFileName))
                    {
                        defaultFileNames.Add(selectedFileName);
                    }

                    // 設定を保存
                    SaveSettings();
                }
            }
        }

        private bool DeleteImageFile(string filePath)
        {
            // 一致する ImageFile を検索
            var itemToRemove = ImageFiles.FirstOrDefault(item => item.FileName == Path.GetFileName(filePath));

            if (itemToRemove != null)
            {
                // コレクションから削除
                ImageFiles.Remove(itemToRemove);

                // 実際のファイルも削除する場合は、以下のコードを追加
                if (File.Exists(filePath))
                {
                    // ファイルがロックされているかどうかを確認
                    if (!IsFileLocked(filePath))
                    {
                        try
                        {
                            // ファイルをゴミ箱に移動
                            FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"ファイルの削除に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("ファイルがロックされているため、削除できません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            return false;
        }

        private string SaveBitmapAsPng(Bitmap bitmap, string fileName)
        {
            if (!Directory.Exists(saveFolderPath))
            {
                Directory.CreateDirectory(saveFolderPath);
            }

            // フルファイル名を作成
            string fullFileName = $"{fileName}.png";
            string filePath = Path.Combine(saveFolderPath, fullFileName);

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
                    return null;
                }

                // リストから削除
                if (DeleteImageFile(filePath) == false)
                {
                    // 上書きをキャンセル
                    return null;
                }
            }

            // PNGとして保存
            bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

            return filePath;
        }

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

        private BitmapImage ConvertBitmapSourceToBitmapImage(BitmapSource bitmapSource)
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

        private void AddImageToList(string filePath)
        {
            try
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

                    // xamlでImageを記述 → imgSync
                    // 画像ファイル情報をリストに追加
                    ImageFiles.Add(new ImageFile
                    {
                        FileName = Path.GetFileName(filePath),
                        Thumbnail = ConvertBitmapSourceToBitmapImage(bmp)
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"画像の読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    saveFolderPath = dialog.SelectedPath;
                    SelectedFolderPath.Text = "選択されたフォルダ: " + saveFolderPath;

                    // 選択されたフォルダの一覧を表示
                    LoadImagesFromFolder();

                    // 設定を保存
                    SaveSettings();
                }
            }
        }

        // サムネイルをクリックしたときの処理
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

        private bool IsFileLocked(string filePath)
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

        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 選択されたサムネイルを取得
            var menuItem = sender as MenuItem;
            var imageFile = menuItem.DataContext as ImageFile;

            if (imageFile == null) return;

            string filePath = Path.Combine(saveFolderPath, imageFile.FileName);

            if (IsFileLocked(filePath))
            {
                MessageBox.Show("ファイルがロックされているため、削除できません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 確認ダイアログを表示
            var result = MessageBox.Show(
                $"このファイルを削除しますか？\n{filePath}",
                "削除確認",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // ファイルをゴミ箱に移動
                    FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);

                    // リストから削除
                    ImageFiles.Remove(imageFile);

                    MessageBox.Show("ファイルはゴミ箱に移動されました。");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ファイルの削除に失敗しました: " + ex.Message);
                }
            }
        }

        // ウィンドウの位置とサイズを保存するため、ウィンドウが閉じられた時に呼び出す
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            SaveSettings();
        }
    }

    // 画像ファイル情報を保持するクラス
    public class ImageFile
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public BitmapImage Thumbnail { get; set; }
    }

    // 設定を保持するクラス
    [Serializable]
    public class AppSettings
    {
        public string SaveFolderPath { get; set; }
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public double WindowTop { get; set; }
        public double WindowLeft { get; set; }
        public double WindowWidth { get; set; }
        public double WindowHeight { get; set; }
    }
}