using ScreenCaptureTool.Models;
using ScreenCaptureTool.Models.CaptureItem;

using System.Windows;

using static ScreenCaptureTool.Models.RecordingSettings;

namespace ScreenCaptureTool.Windows
{
    /// <summary>
    /// CaptureSettingsWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class CaptureSettingsWindow : Window
    {
        #region Properties

        public RecordingSettings Settings { get; private set; }

        #endregion Properties

        #region Constructors

        public CaptureSettingsWindow()
        {
            InitializeComponent();

            // 新規用に設定を作成
            Settings = new RecordingSettings("新規設定");
        }

        #endregion Constructors

        #region Methods

        #region Methods(Public)

        /// <summary>
        /// 設定の読み込み（画面UIへの反映）
        /// </summary>
        /// <param name="setting">撮影設定</param>
        public void LoadSettings(RecordingSettings setting)
        {
            // 保持
            Settings = setting;

            // 名称
            textBoxName.Text = setting.Name;

            // 撮影方法：ウィンドウ
            radioButtonWindow.IsChecked = setting.Type == RecordingSettings.RecordingType.Window;

            // ウィンドウタイトル
            textBoxWindowTitle.Text = setting.WindowTitle;

            // 撮影方法：画面の一部
            radioButtonRectangle.IsChecked = setting.Type == RecordingSettings.RecordingType.ScreenRect;

            // 撮影位置固定
            checkBoxFixedPos.IsChecked = setting.LocationFixed;
            textBoxWindowPosX.Text = setting.Location.X.ToString();
            textBoxWindowPosY.Text = setting.Location.Y.ToString();

            // 撮影サイズ固定
            checkBoxFixedSize.IsChecked = setting.SizeFixed;
            textBoxWindowWidth.Text = setting.Size.Width.ToString();
            textBoxWindowHeight.Text = setting.Size.Height.ToString();

            // 縁トリミング
            checkBoxTrimEdge.IsChecked = setting.TrimEdgeEnabled;
            textBoxTrimTop.Text = setting.TrimEdgeInset.Top.ToString();
            textBoxTrimBottom.Text = setting.TrimEdgeInset.Bottom.ToString();
            textBoxTrimLeft.Text = setting.TrimEdgeInset.Left.ToString();
            textBoxTrimRight.Text = setting.TrimEdgeInset.Right.ToString();

            // リサイズ
            checkBoxResize.IsChecked = setting.ResizeEnabled;
            textBoxResizeWidth.Text = setting.ResizeSize.Width.ToString();
            textBoxResizeHeight.Text = setting.ResizeSize.Height.ToString();

            // 保存形式
            radioButtonSaveClipboard.IsChecked = setting.SaveType == RecordingSettings.ImageSaveType.Clipboard;
            radioButtonSavePNG.IsChecked = setting.SaveType == RecordingSettings.ImageSaveType.FilePNG;
            radioButtonSaveBMP.IsChecked = setting.SaveType == RecordingSettings.ImageSaveType.FileBMP;
            radioButtonSaveJPEG.IsChecked = setting.SaveType == RecordingSettings.ImageSaveType.FileJPEG;

            // UIコントロールの更新
            RefreshUIControls();
        }

        #endregion Methods(Public)

        #region Methods(Private)

        /// <summary>
        /// 現在の画面UIの設定値から、撮影設定の取得
        /// </summary>
        /// <returns>撮影設定</returns>
        private RecordingSettings SaveSettings()
        {
            // 名称
            Settings.Name = textBoxName.Text;

            // 撮影方法：ウィンドウ
            if (radioButtonWindow.IsChecked == true)
            {
                Settings.Type = RecordingSettings.RecordingType.Window;
                Settings.WindowTitle = textBoxWindowTitle.Text;
            }

            // 撮影方法：画面の一部
            if (radioButtonRectangle.IsChecked == true)
            {
                Settings.Type = RecordingSettings.RecordingType.ScreenRect;
                Settings.LocationFixed = checkBoxFixedPos.IsChecked ?? false;
                Settings.Location = new System.Drawing.Point(int.Parse(textBoxWindowPosX.Text), int.Parse(textBoxWindowPosY.Text));
                Settings.SizeFixed = checkBoxFixedSize.IsChecked ?? false;
                Settings.Size = new System.Drawing.Size(int.Parse(textBoxWindowWidth.Text), int.Parse(textBoxWindowHeight.Text));
            }

            // 縁トリミング
            Settings.TrimEdgeEnabled = checkBoxTrimEdge.IsChecked ?? false;
            Settings.TrimEdgeInset = new EdgeInsets(int.Parse(textBoxTrimTop.Text), int.Parse(textBoxTrimBottom.Text), int.Parse(textBoxTrimLeft.Text), int.Parse(textBoxTrimRight.Text));

            // リサイズ
            Settings.ResizeEnabled = checkBoxResize.IsChecked ?? false;
            Settings.ResizeSize = new System.Drawing.Size(int.Parse(textBoxResizeWidth.Text), int.Parse(textBoxResizeHeight.Text));

            // 保存形式
            if (radioButtonSaveClipboard.IsChecked == true) Settings.SaveType = RecordingSettings.ImageSaveType.Clipboard;
            if (radioButtonSavePNG.IsChecked == true) Settings.SaveType = RecordingSettings.ImageSaveType.FilePNG;
            if (radioButtonSaveBMP.IsChecked == true) Settings.SaveType = RecordingSettings.ImageSaveType.FileBMP;
            if (radioButtonSaveJPEG.IsChecked == true) Settings.SaveType = RecordingSettings.ImageSaveType.FileJPEG;

            return Settings;
        }

        /// <summary>
        /// UIコントロールの更新
        /// </summary>
        private void RefreshUIControls()
        {
            // 撮影方法：ウィンドウ
            if (radioButtonWindow.IsChecked == true)
            {
                RecordingTypeWindowPanel.IsEnabled = true;
                RecordingTypeRectanglePanel.IsEnabled = false;
            }

            // 撮影方法：画面の一部
            if (radioButtonRectangle.IsChecked == true)
            {
                RecordingTypeWindowPanel.IsEnabled = false;
                RecordingTypeRectanglePanel.IsEnabled = true;

                recordingPositionPanel.IsEnabled = checkBoxFixedPos.IsChecked ?? false;
                recordingSizePanel.IsEnabled = checkBoxFixedSize.IsChecked ?? false;
            }

            // 加工
            trimEdgePanel.IsEnabled = checkBoxTrimEdge.IsChecked ?? false;
            resizePanel.IsEnabled = checkBoxResize.IsChecked ?? false;
        }

        #region Methods(Event)

        /// <summary>
        /// UIコントロールの更新
        /// </summary>
        private void RefreshUIControls(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// OKボタン：押下イベント
        /// </summary>
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            // 変更内容を反映
            SaveSettings();
            
            DialogResult = true;          // ダイアログを閉じて結果を返す
            Close();
        }

        /// <summary>
        /// キャンセルボタン：押下イベント
        /// </summary>
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;         // ダイアログを閉じて結果を返す
            Close();
        }

        #endregion Methods(Event)

        #endregion Methods(Private)

        #endregion Methods
    }
}