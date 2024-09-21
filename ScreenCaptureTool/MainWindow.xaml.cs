using System;
using System.IO;
using System.Drawing;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Xml.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.VisualBasic.FileIO;

using PixelFormat = System.Drawing.Imaging.PixelFormat;
using MessageBox = System.Windows.MessageBox;

namespace ScreenCaptureTool
{
    public partial class MainWindow : Window
    {
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
        /// アプリケーション設定を保存する XML ファイル名
        /// </summary>
        private const string SettingsFilePath = "settings.xml";

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

            // ComboBox にファイル名リストを設定
            FileNameComboBox.ItemsSource = SaveFileNames;

            // 設定を読み込む
            LoadSettings();

            // フォルダツリーの初期化
            InitializeFolderTree();

            // 保存先フォルダをツリーから選択状態にする
            SelectFolderInTree(saveFolderPath);
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
            SaveSettings();
        }

        #endregion Events(Window)

        #region Private

        #region Settings

        /// <summary>
        /// アプリケーション設定をXMLファイルから読み込む
        /// </summary>
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

                        // 保存ファイル名一覧
                        SaveFileNames = settings.SaveFileNames;
                        FileNameComboBox.ItemsSource = SaveFileNames;

                        // 保存先フォルダ
                        saveFolderPath = settings.SaveFolderPath;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("設定ファイルの読み込みに失敗しました: " + ex.Message);
                }
            }

            // UIに反映
            SelectedFolderTextBox.Text = saveFolderPath;
        }

        /// <summary>
        /// アプリケーション設定をXMLファイルに保存する
        /// </summary>
        private void SaveSettings()
        {
            try
            {
                AppSettings settings = new AppSettings();
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
                // サムネイルサイズ
                settings.ThumbnailSize = thumbnailSize;
                // 保存ファイル名一覧
                settings.SaveFileNames = SaveFileNames;
                // 保存先フォルダ
                settings.SaveFolderPath = saveFolderPath;

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

        #region ScreenCapture

        /// <summary>
        /// デスクトップの指定範囲をキャプチャーする
        /// </summary>
        private void CaptureDesktop()
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

            CaptureDesktop(x, y, width, height);
        }

        /// <summary>
        /// デスクトップの指定範囲をキャプチャーする
        /// </summary>
        /// <param name="x">X</param>
        /// <param name="y">Y</param>
        /// <param name="width">幅</param>
        /// <param name="height">高さ</param>
        private void CaptureDesktop(int x, int y, int width, int height)
        {
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
                string filePath = SaveBitmapAsPng(bitmap, saveFolderPath, selectedFileName);

                if (filePath != null)
                {
                    // サムネイルリストを更新
                    AddImageToList(filePath);

                    // 新しいファイル名をComboBoxのリストに追加
                    if (!FileNameComboBox.Items.Contains(selectedFileName))
                    {
                        SaveFileNames.Add(selectedFileName);
                    }

                    // 設定を保存
                    SaveSettings();
                }
            }
        }

        #endregion ScreenCapture

        #region Thumbnails

        /// <summary>
        /// BitmapをPNG形式で保存する
        /// </summary>
        /// <param name="bitmap">Bitmap</param>
        /// <param name="folderPath">フォルダパス</param>
        /// <param name="fileName">ファイル名(拡張子は含まない)</param>
        /// <returns>保存先のパス(失敗時はnull)</returns>
        private string SaveBitmapAsPng(Bitmap bitmap, string folderPath, string fileName)
        {
            // パスを作成
            string fullFileName = $"{fileName}.png";
            string filePath = Path.Combine(folderPath, fullFileName);

            // フォルダが存在しない場合は作成する
            if (!Directory.Exists(Path.GetDirectoryName(filePath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(filePath));
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

        /// <summary>
        /// 画像ファイルを削除(ファイルはゴミ箱に移動、画像一覧からも削除)
        /// </summary>
        /// <param name="filePath">パス</param>
        /// <returns>true:成功 / false:失敗</returns>
        private bool DeleteImageFile(string filePath)
        {
            // 実際のファイルも削除
            if (File.Exists(filePath))
            {
                // ファイルがロックされているかどうかを確認
                if (!IsFileLocked(filePath))
                {
                    try
                    {
                        // ファイルをゴミ箱に移動
                        FileSystem.DeleteFile(filePath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ファイルの削除に失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("ファイルがロックされているため、削除できません。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show($"画像の読み込みに失敗しました: {ex.Message}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
                SelectCurrentFolder(selectedItem.Tag.ToString());
            }
        }

        // saveFolderPathのフォルダをツリーで選択する
        private void SelectFolderInTree(string saveFolderPath)
        {
            string[] pathParts = saveFolderPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            TreeViewItem currentItem = null;

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
                return;

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

            // 設定を保存
            SaveSettings();

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
                    var directories = Directory.GetDirectories(selectedItem.Tag.ToString());
                    foreach (var directory in directories)
                    {
                        var subItem = new TreeViewItem { Header = Path.GetFileName(directory), Tag = directory };
                        subItem.Items.Add(null);  // ダミーアイテム
                        subItem.Expanded += Folder_Expanded;
                        selectedItem.Items.Add(subItem);
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    // アクセスできないフォルダに対してのエラーハンドリング
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
        /// キャプチャーボタン：押下
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CaptureDesktopButton_Click(object sender, RoutedEventArgs e)
        {
            // デスクトップ全体をキャプチャ
            CaptureDesktop();
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

        #endregion Events(Control)

        #endregion Methods
    }
}