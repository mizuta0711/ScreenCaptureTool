using System;
using System.IO;
using System.Drawing;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.VisualBasic.FileIO;
using MessageBox = System.Windows.MessageBox;
using ScreenCaptureTool.Models;
using ScreenCaptureTool.Models.CaptureItem;
using ScreenCaptureTool.Utilities;
using System.Windows.Input;
using ScreenCaptureTool.Controllers.CaptureProvider;

namespace ScreenCaptureTool.Windows
{
    public partial class MainWindow : Window
    {
        #region Variable

        /// <summary>
        /// 撮影設定一覧
        /// </summary>
        public ObservableCollection<RecordingSetting> RecordingSettings { get; set; } = new ObservableCollection<RecordingSetting>();

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
        private ProjectSetting projectSettings = new ProjectSetting();

        private RecordingSetting? recordingSetting;

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

            // プロジェクトファイルを読み込む
            LoadProjectFile(projectSettings.FilePath);
        }

        #endregion Constructor

        #region Methods

        #region Events(Window)

        /// <summary>
        /// ウィンドウが閉じられた：ウィンドウの位置とサイズを保存する
        /// </summary>
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
        /// <param name="setting">プロジェクト設定</param>
        private void LoadProjectSettings(ProjectSetting setting)
        {
            // サムネイルサイズ
            if (setting.ThumbnailSize > 0)
            {
                thumbnailSize = setting.ThumbnailSize;
            }

            // ウィンドウの位置とサイズを設定
            if (setting.WindowTop >= 0 && setting.WindowLeft >= 0)
            {
                Top = setting.WindowTop;
                Left = setting.WindowLeft;
            }
            if (setting.WindowWidth > 0 && setting.WindowHeight > 0)
            {
                Width = setting.WindowWidth;
                Height = setting.WindowHeight;
            }

            // 保存先フォルダ
            saveFolderPath = setting.SaveFolderPath;

            // 保存先フォルダをツリーから選択状態にする
            FolderTreeView.SelectFolderInTree(saveFolderPath);

            // サムネイルサイズメニューのチェックを設定
            var menuItems = ThumbnailSizeMenu.Items.Cast<MenuItem>();
            var defaultItem = menuItems.FirstOrDefault(item => int.Parse((string)item.Tag) == thumbnailSize);
            if (defaultItem != null)
            {
                defaultItem.IsChecked = true;
            }

            // 画像一覧を読み込む
            RecordingSettings = setting.RecordingSettings;
        }

        /// <summary>
        /// UIの設定をプロジェクト設定に保存する
        /// </summary>
        /// <param name="setting">プロジェクト設定</param>
        private void StoreProjectSettings(ProjectSetting setting)
        {
            // ウィンドウの位置とサイズ
            setting.WindowTop = Top;
            setting.WindowLeft = Left;
            setting.WindowWidth = Width;
            setting.WindowHeight = Height;
            // サムネイルサイズ
            setting.ThumbnailSize = thumbnailSize;
            // 保存先フォルダ
            setting.SaveFolderPath = saveFolderPath;
            // 撮影設定
            setting.RecordingSettings = RecordingSettings;
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

            var settings = ProjectSetting.Load(filePath);
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
        private bool SaveProjectFile(ProjectSetting settings)
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
            if (recordingSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return false;
            }

            if (recordingSetting.SaveType == RecordingSetting.ImageSaveType.Clipboard)
            {
                // クリップボードにコピー
                BitmapHelper.CopyToClipboard(bitmap);
                return true;
            }

            // ComboBoxから選択または入力されたファイル名を取得
            string selectedFileName = recordingSetting.Name.Trim();
            if (string.IsNullOrEmpty(selectedFileName))
            {
                ShowErrorDialog("ファイル名を入力してください。");
                return false;
            }

            try
            {
                // 画像ファイルに保存
                string filePath = CreateFilePath(saveFolderPath, selectedFileName, recordingSetting.FileExtension);
                if (BitmapHelper.SaveToFile(bitmap, filePath, recordingSetting.SaveType, recordingSetting.ConfirmOverrideFile) == false)
                {
                    return false;
                }

                // サムネイルリストを更新
                AddImageToList(filePath);

                return true;
            }
            catch (Exception ex)
            {
                ShowErrorDialog("画像の保存に失敗しました: " + ex.Message);
                return false;
            }
        }

        #endregion CaptureTools

        #region Thumbnails

        /// <summary>
        /// 保存先フォルダの画像一覧を取得する
        /// </summary>
        private void LoadImagesFromFolder()
        {
            ImageFiles.Clear();

            // カーソルを待機中に変更
            Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            // TODO: 非同期処理にリファクタリングする
            if (Directory.Exists(saveFolderPath))
            {
                var files = Directory.GetFiles(saveFolderPath, "*.png");
                foreach (var file in files)
                {
                    AddImageToList(file);
                }
            }

            // カーソルを元に戻す
            Mouse.OverrideCursor = null;
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
                BitmapImage thumbnail = BitmapHelper.LoadBitmapImage(filePath);

                // 一覧に同じファイル名があれば更新、なければ追加
                foreach (var imageFile in ImageFiles)
                {
                    if (imageFile.FileName == Path.GetFileName(filePath))
                    {
                        // 画像ファイル情報を更新
                        imageFile.Thumbnail = thumbnail;
                        imageFile.ThumbnailWidth = thumbnailSize;
                        imageFile.ThumbnailHeight = thumbnailSize;
                        return;
                    }
                }

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

        /// <summary>
        /// サムネイル画像を選択状態にする
        /// </summary>
        /// <param name="selectedImageFile">選択したImageFile</param>
        private void SelectImageFile(ImageFile selectedImageFile)
        {
            // 他のアイテムをすべて非選択にする
            foreach (var item in ThumbnailItemsControl.ItemsSource)
            {
                if (item is ImageFile imageFile)
                {
                    imageFile.IsSelected = false;
                }
            }

            // クリックされたアイテムを選択状態にする
            selectedImageFile.IsSelected = true;
        }

        /// <summary>
        /// サムネイル画像を開く
        /// </summary>
        /// <param name="imageFile">サムネイル画像</param>
        private void OpenThumbnailImage(ImageFile imageFile)
        {
            // フルパスをProcess.Startに渡す
            string fullPath = Path.Combine(saveFolderPath, imageFile.FileName ?? "");

            // フォトアプリで画像を開く
            Process.Start(new ProcessStartInfo(fullPath)
            {
                UseShellExecute = true // 既定のアプリケーションで開く
            });
        }

        #endregion Thumbnails

        #region FolderTreeView

        /// <summary>
        /// カレントフォルダの変更
        /// </summary>
        /// <param name="folderPath">パス</param>
        private void SelectCurrentFolder(string folderPath)
        {
            // 保存先フォルダを変更
            saveFolderPath = folderPath;

            // 選択されたフォルダの一覧を表示
            LoadImagesFromFolder();

            // 選択されたフォルダのツリー更新
            FolderTreeView.RefreshSelectedFolderTree();
        }

        #endregion FolderTreeView

        #endregion Private

        #region Events(Menu)

        /// <summary>
        /// 開く：メニュー
        /// </summary>
        private void OnOpenProjectMenu_Clicked(object sender, RoutedEventArgs e)
        {
            // ダイアログでファイルを選択
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Screen Capture Project (*.scp)|*.scp";
            if (openFileDialog.ShowDialog() == true)
            {
                LoadProjectFile(openFileDialog.FileName);
            }
        }

        /// <summary>
        /// 上書き保存：メニュー
        /// </summary>
        private void OnSaveProjectMenu_Clicked(object sender, RoutedEventArgs e)
        {
            if (projectSettings.FilePath != null)
            {
                SaveProjectFile(projectSettings);
            }
            else
            {
                OnSaveAsProjectMenu_Clicked(sender, e);
            }
        }

        /// <summary>
        /// 名前を付けて保存：メニュー
        /// </summary>
        private void OnSaveAsProjectMenu_Clicked(object sender, RoutedEventArgs e)
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

        /// <summary>
        /// 終了：メニュー
        /// </summary>
        private void OnExitMenu_Clicked(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// 表示：サムネイルサイズメニュー
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThumbnailSizeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 全てのサイズメニューのチェックを解除
            foreach (var menuItem in ((MenuItem)((MenuItem)sender).Parent).Items)
            {
                if (menuItem is MenuItem item)
                {
                    item.IsChecked = false;
                }
            }

            // クリックされた項目をチェックし、サムネイルサイズを設定
            var selectedItem = sender as MenuItem;
            if (selectedItem != null)
            {
                selectedItem.IsChecked = true;
                // メニューのTagからサイズを取得
                thumbnailSize = int.Parse((string)selectedItem.Tag);
                UpdateThumbnailsSize();  // サイズ変更後にサムネイル更新
            }
        }

        #endregion Events(Menu)

        #region Events(FolderTreeView)

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

        #endregion Events(FolderTreeView)

        #region Events(Control)

        /// <summary>
        /// フォルダ選択ボタン：押下
        /// </summary>
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
        /// 設定追加ボタン：押下
        /// </summary>
        private void AddSettingButton_Click(object sender, RoutedEventArgs e)
        {
            RecordingSettingsWindow dialog = new RecordingSettingsWindow();
            dialog.Owner = this;    // 現在のウィンドウを親に設定
            if (dialog.ShowDialog() == true)
            {
                RecordingSettings.Add(dialog.Setting);
            }
        }

        /// <summary>
        /// 設定変更ボタン：押下
        /// </summary>
        private void EditSettingButton_Click(object sender, RoutedEventArgs e)
        {
            if (recordingSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return;
            }

            RecordingSettingsWindow dialog = new RecordingSettingsWindow();
            dialog.Owner = this;    // 現在のウィンドウを親に設定
            dialog.LoadSettings(recordingSetting);
            if (dialog.ShowDialog() == true)
            {
            }
        }

        /// <summary>
        /// 撮影ボタン：押下
        /// </summary>
        private void CaptureButton_Click(object sender, RoutedEventArgs e)
        {
            if (recordingSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return;
            }

            try
            {
                // ファイル名が未入力の場合はエラー
                if (recordingSetting.SaveType != RecordingSetting.ImageSaveType.Clipboard &&
                    string.IsNullOrEmpty(recordingSetting.Name))
                {
                    ShowErrorDialog("ファイル名を入力してください。");
                    return;
                }

                var captureItem = recordingSetting.GetCaptureItem();
                // キャプチャ処理
                var provider = CaptureProviderFactory.InstantiateCaptureProvider(captureItem);
                var bitmap = provider.Capture();
                if (bitmap != null)
                {
                    if (provider.CaptureItem is ManualScreenRectCaptureItem manualScreenRectCaptureItem)
                    {
                        // 選択された矩形を取得
                        recordingSetting.Location = manualScreenRectCaptureItem.TargetRect.Location;
                        recordingSetting.Size = manualScreenRectCaptureItem.TargetRect.Size;
                    }

                    SaveCaptureImage(bitmap);
                }
            }
            catch (ArgumentException ex)
            {
                Debug.WriteLine(ex.Message);
                ShowErrorDialog("撮影設定に誤りがあります");
                // 撮影設定を開く
                EditSettingButton_Click(sender, e);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                ShowErrorDialog("撮影に失敗しました\n" + ex.Message);
            }
        }

        /// <summary>
        /// サムネイル画像：左クリック
        /// </summary>
        private void Thumbnail_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // ダブルクリック：画像を開く
            if (sender is System.Windows.Controls.Image image)
            {
                if (image.DataContext is ImageFile clickedItem)
                {
                    switch (e.ClickCount)
                    {
                        case 1:
                            // サムネイル画像を選択
                            SelectImageFile(clickedItem);
                            break;

                        case 2:
                            // ダブルクリック：画像を開く
                            OpenThumbnailImage(clickedItem);
                            break;

                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 削除メニュー：選択
        /// </summary>
        private void DeleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // 選択されたサムネイルを取得
            var menuItem = sender as MenuItem;
            if (menuItem == null) return;
            var imageFile = menuItem.DataContext as ImageFile;

            if (imageFile == null) return;
            if (imageFile.FileName == null) return;

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

                    // ShowInformationDialog("ファイルはゴミ箱に移動されました。");
                }
                catch (Exception ex)
                {
                    ShowErrorDialog("ファイルの削除に失敗しました: " + ex.Message);
                }
            }
        }

        private void SettingListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 選択されたアイテムを取得
            if (sender is System.Windows.Controls.ListView listView)
            {
                if (listView.SelectedItem is RecordingSetting setting)
                {
                    // 選択された設定を保持
                    recordingSetting = setting;
                }
            }
        }

        #endregion Events(Control)

        #endregion Methods
    }
}