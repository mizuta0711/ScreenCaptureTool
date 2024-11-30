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
using System.Windows.Documents;

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

        private RecordingSetting? CurrentSetting
        {
            get
            {
                if (listViewRecoringSetting.SelectedItem is RecordingSetting selectedSetting)
                {
                    return selectedSetting;
                }
                return null;
            }
        }

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

        #region Methods(Override)

        /// <summary>
        /// ウィンドウが閉じられた：ウィンドウの位置とサイズを保存する
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // プロジェクトファイルを保存
            SaveProjectFile(projectSettings);
        }

        #endregion Methods(Override)

        #region Methods(Private)

        #region Methods(Utility)

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

        #endregion Methods(Utility)

        #region Methods(ProjectSettings)

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

        #endregion Methods(ProjectSettings)

        #region Methods(RecordingSettings)

        /// <summary>
        /// 撮影設定一覧の右クリックメニューが開かれるとき
        /// </summary>
        private void RecordingSettingListViewMenu_Opened(object sender, RoutedEventArgs e)
        {
            // 現在選択されている項目があるかをチェック
            bool hasSelectedItem = listViewRecoringSetting.SelectedItem != null;

            // 「編集」と「削除」メニューの有効/無効を切り替え
            MenuItemEdit.IsEnabled = hasSelectedItem;
            MenuItemDelete.IsEnabled = hasSelectedItem;
        }

        /// <summary>
        /// 撮影設定の追加
        /// </summary>
        private void AddRecordingSetting()
        {
            RecordingSettingsWindow dialog = new RecordingSettingsWindow();
            dialog.Owner = this;    // 現在のウィンドウを親に設定
            if (dialog.ShowDialog() == true)
            {
                RecordingSettings.Add(dialog.Setting);
            }
        }

        /// <summary>
        /// 撮影設定の編集
        /// </summary>
        /// <param name="setting">撮影設定</param>
        /// <returns>true: 編集した / false: キャンセル</returns>
        private bool EditRecordingSetting(RecordingSetting setting)
        {
            RecordingSettingsWindow dialog = new RecordingSettingsWindow();
            dialog.Owner = this;    // 現在のウィンドウを親に設定
            dialog.LoadSettings(setting);
            if (dialog.ShowDialog() == true)
            {
                var view = System.Windows.Data.CollectionViewSource.GetDefaultView(RecordingSettings);
                view.Refresh();
                return true;
            }
            return false;
        }

        /// <summary>
        /// 撮影設定の削除
        /// </summary>
        /// <param name="setting">撮影設定</param>
        /// <returns>true: 削除した / false: キャンセル</returns>
        private bool DeleteRecordingSetting(RecordingSetting setting)
        {
            // 確認ダイアログを表示
            if (MessageBox.Show($"設定「{setting.Name}」を削除しますか？", "削除確認", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                RecordingSettings.Remove(setting);
                return true;
            }
            return false;
        }

        #endregion Methods(RecordingSettings)

        #region Methods(CaptureTools)

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
            if (CurrentSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return false;
            }

            if (CurrentSetting.SaveType == RecordingSetting.ImageSaveType.Clipboard)
            {
                // クリップボードにコピー
                BitmapHelper.CopyToClipboard(bitmap);
                return true;
            }

            // ファイル名を取得
            var saveFileFullPath = Path.Combine(saveFolderPath, CurrentSetting.SaveFileName);
            var saveType = CurrentSetting.SaveType;
            if (CurrentSetting.SaveAsEnable)
            {
                // 名前を付けて保存ダイアログを表示
                Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                saveFileDialog.Filter = "PNGファイル (*.png)|*.png|BMPファイル (*.bmp)|*.bmp|JPEGファイル (*.jpg)|*.jpg|全てのファイル (*.*)|*.*";
                saveFileDialog.InitialDirectory = Path.GetDirectoryName(saveFileFullPath);
                saveFileDialog.FileName = Path.GetFileName(saveFileFullPath);
                saveFileDialog.OverwritePrompt = false;
                if (saveFileDialog.ShowDialog() == false)
                {
                    ShowErrorDialog("キャンセルしました");
                    return false;
                }

                // ファイル名を取得
                saveFileFullPath = saveFileDialog.FileName;
                // 拡張子に応じた画像形式に変更
                saveType = RecordingSetting.GetImageSaveType(saveFileFullPath);
            }

            try
            {
                // 画像ファイルに保存
                if (BitmapHelper.SaveToFile(bitmap, saveFileFullPath, saveType, CurrentSetting.ConfirmOverrideFile) == false)
                {
                    return false;
                }

                // サムネイルリストを更新
                AddImageToList(saveFileFullPath);

                return true;
            }
            catch (Exception ex)
            {
                ShowErrorDialog("画像の保存に失敗しました: " + ex.Message);
                return false;
            }
        }

        #endregion Methods(CaptureTools)

        #region Methods(Thumbnails)

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
                // 一覧に同じファイル名があれば更新、なければ追加
                foreach (var imageFile in ImageFiles)
                {
                    if (imageFile.FileName == Path.GetFileName(filePath))
                    {
                        // 画像ファイル情報を更新
                        imageFile.ImageUrl = filePath;
                        //imageFile.Thumbnail = thumbnail;
                        imageFile.ThumbnailWidth = thumbnailSize;
                        imageFile.ThumbnailHeight = thumbnailSize;
                        return;
                    }
                }

                // 画像ファイル情報をリストに追加
                ImageFiles.Add(new ImageFile
                {
                    FileName = Path.GetFileName(filePath),
                    ImageUrl = filePath,
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

        #endregion Methods(Thumbnails)

        #region Methods(FolderTreeView)

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

        #endregion Methods(FolderTreeView)

        #region Methods(Event)

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

        #region Events(RecordingSettingListView)

        /// <summary>
        /// 撮影設定一覧：左ダブルクリック
        /// </summary>
        private void RecordingSettingList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (CurrentSetting != null)
            {
                EditRecordingSetting(CurrentSetting);
            }
        }

        /// <summary>
        /// 撮影設定一覧のコンテキストメニュー：追加
        /// </summary>
        private void AddSettingMenuItem_Click(object sender, RoutedEventArgs e)
        {
            AddRecordingSetting();
        }

        /// <summary>
        /// 撮影設定一覧のコンテキストメニュー：編集
        /// </summary>
        private void EditSettingMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (listViewRecoringSetting.SelectedItem is RecordingSetting selectedSetting)
            {
                EditRecordingSetting(selectedSetting);
            }
        }

        /// <summary>
        /// 撮影設定一覧のコンテキストメニュー：削除
        /// </summary>
        private void DeleteSettingMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentSetting != null)
            {
                DeleteRecordingSetting(CurrentSetting);
            }
        }

        #endregion Events(RecordingSettingListView)

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
            AddRecordingSetting();
        }

        /// <summary>
        /// 設定変更ボタン：押下
        /// </summary>
        private void EditSettingButton_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return;
            }

            EditRecordingSetting(CurrentSetting);
        }

        /// <summary>
        /// 撮影ボタン：押下
        /// </summary>
        private void CaptureButton_Click(object sender, RoutedEventArgs e)
        {
            if (CurrentSetting == null)
            {
                ShowErrorDialog("撮影設定が選択されていません。");
                return;
            }

            try
            {
                // ファイル名が未入力の場合はエラー
                if (CurrentSetting.SaveType != RecordingSetting.ImageSaveType.Clipboard &&
                    string.IsNullOrEmpty(CurrentSetting.Name))
                {
                    ShowErrorDialog("ファイル名を入力してください。");
                    return;
                }

                var captureItem = CurrentSetting.GetCaptureItem();
                // キャプチャ処理
                var provider = CaptureProviderFactory.InstantiateCaptureProvider(captureItem);
                var bitmap = provider.Capture();
                if (bitmap != null)
                {
                    // 画像を保存
                    if (SaveCaptureImage(bitmap))
                    {
                        // 保存が成功したら撮影設定を更新して保存
                        // 撮影位置とサイズが固定されていない場合は、撮影後の位置とサイズを保存
                        if (provider.CaptureItem is ManualScreenRectCaptureItem manualScreenRectCaptureItem)
                        {
                            // 選択された矩形を取得
                            if (CurrentSetting.LocationFixed == false)
                            {
                                CurrentSetting.Location = manualScreenRectCaptureItem.TargetRect.Location;
                            }
                            if (CurrentSetting.SizeFixed == false)
                            {
                                CurrentSetting.Size = manualScreenRectCaptureItem.TargetRect.Size;
                            }
                        }
                        // ファイル名書式のインクリメント
                        CurrentSetting.IncrimentFilenameFormatNumber();
                        // 撮影設定を保存
                        SaveProjectFile(projectSettings);
                    }
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
        /// サムネイル画像：右クリック
        /// </summary>
        private void Thumbnail_MouseRightButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            // シングルクリック時の選択処理
            if (e.ClickCount == 1)
            {
                // サムネイル画像を選択
                if (sender is System.Windows.Controls.Image image)
                {
                    if (image.DataContext is ImageFile clickedItem)
                    {
                        SelectImageFile(clickedItem);
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
            DeleteImageFile(filePath, true);
        }

        #endregion Events(Control)

        #endregion Methods(Event)

        #endregion Methods(Private)

        #endregion Methods
    }
}