using ScreenCaptureTool.Models;
using ScreenCaptureTool.Models.CaptureItem;

using System.Windows;

using static ScreenCaptureTool.Models.RecordingSetting;

namespace ScreenCaptureTool.Windows
{
    /// <summary>
    /// RecordingSettingsWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class RecordingSettingsWindow : Window
    {
        #region Properties

        public RecordingSetting Setting { get; private set; }

        #endregion Properties

        #region Constructors

        public RecordingSettingsWindow()
        {
            InitializeComponent();

            // 新規用に設定を作成
            Setting = new RecordingSetting("新規設定");
            LoadSettings(Setting);
        }

        #endregion Constructors

        #region Methods

        #region Methods(Public)

        /// <summary>
        /// 設定の読み込み（画面UIへの反映）
        /// </summary>
        /// <param name="setting">撮影設定</param>
        public void LoadSettings(RecordingSetting setting)
        {
            // 保持
            Setting = setting;

            // 名称
            textBoxName.Text = setting.Name;

            // 説明
            textBoxDescription.Text = setting.Description;

            // 撮影方法：ウィンドウ
            radioButtonWindow.IsChecked = setting.Type == RecordingSetting.RecordingType.Window;

            // ウィンドウタイトル
            textBoxWindowTitle.Text = setting.WindowTitle;

            // 撮影方法：画面の一部
            radioButtonRectangle.IsChecked = setting.Type == RecordingSetting.RecordingType.ScreenRect;

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
            radioButtonSaveClipboard.IsChecked = setting.SaveType == RecordingSetting.ImageSaveType.Clipboard;
            radioButtonSavePNG.IsChecked = setting.SaveType == RecordingSetting.ImageSaveType.FilePNG;
            radioButtonSaveBMP.IsChecked = setting.SaveType == RecordingSetting.ImageSaveType.FileBMP;
            radioButtonSaveJPEG.IsChecked = setting.SaveType == RecordingSetting.ImageSaveType.FileJPEG;

            // 保存形式：ファイルの上書き確認
            checkBoxConfirmOverrideFile.IsChecked = setting.ConfirmOverrideFile;

            // UIコントロールの更新
            RefreshUIControls();
        }

        #endregion Methods(Public)

        #region Methods(Private)

        /// <summary>
        /// 文字列を数値に変換
        /// </summary>
        /// <param name="text">テキスト</param>
        /// <param name="defaultValue">デフォルト値</param>
        /// <returns>数値(変換できない場合はデフォルト値)</returns>
        private int ParseInt(string text, int defaultValue = 0)
        {
            return int.TryParse(text, out int value) ? value : defaultValue;
        }

        /// <summary>
        /// 現在の画面UIの設定値から、撮影設定の取得
        /// </summary>
        /// <returns>撮影設定</returns>
        private RecordingSetting SaveSettings()
        {
            // 名称
            Setting.Name = textBoxName.Text;

            // 説明
            Setting.Description = textBoxDescription.Text;

            // 撮影方法：ウィンドウ
            if (radioButtonWindow.IsChecked == true)
            {
                Setting.Type = RecordingSetting.RecordingType.Window;
                Setting.WindowTitle = textBoxWindowTitle.Text;
            }

            // 撮影方法：画面の一部
            if (radioButtonRectangle.IsChecked == true)
            {
                Setting.Type = RecordingSetting.RecordingType.ScreenRect;
                Setting.LocationFixed = checkBoxFixedPos.IsChecked ?? false;
                Setting.Location = new System.Drawing.Point(ParseInt(textBoxWindowPosX.Text),
                                                            ParseInt(textBoxWindowPosY.Text));
                Setting.SizeFixed = checkBoxFixedSize.IsChecked ?? false;
                Setting.Size = new System.Drawing.Size(ParseInt(textBoxWindowWidth.Text),
                                                       ParseInt(textBoxWindowHeight.Text));
            }

            // 縁トリミング
            Setting.TrimEdgeEnabled = checkBoxTrimEdge.IsChecked ?? false;
            Setting.TrimEdgeInset = new EdgeInsets(ParseInt(textBoxTrimTop.Text),
                                                   ParseInt(textBoxTrimLeft.Text),
                                                   ParseInt(textBoxTrimBottom.Text),
                                                   ParseInt(textBoxTrimRight.Text));

            // リサイズ
            Setting.ResizeEnabled = checkBoxResize.IsChecked ?? false;
            Setting.ResizeSize = new System.Drawing.Size(ParseInt(textBoxResizeWidth.Text),
                                                         ParseInt(textBoxResizeHeight.Text));

            // 保存形式
            if (radioButtonSaveClipboard.IsChecked == true) Setting.SaveType = RecordingSetting.ImageSaveType.Clipboard;
            if (radioButtonSavePNG.IsChecked == true) Setting.SaveType = RecordingSetting.ImageSaveType.FilePNG;
            if (radioButtonSaveBMP.IsChecked == true) Setting.SaveType = RecordingSetting.ImageSaveType.FileBMP;
            if (radioButtonSaveJPEG.IsChecked == true) Setting.SaveType = RecordingSetting.ImageSaveType.FileJPEG;

            // 保存形式：ファイルの上書き確認
            Setting.ConfirmOverrideFile = checkBoxConfirmOverrideFile.IsChecked ?? false;

            return Setting;
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

            // 上書き確認
            checkBoxConfirmOverrideFile.IsEnabled = (radioButtonSaveClipboard.IsChecked ?? true) ? false : true;
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
            // 名称の入力チェック
            textBoxName.Text = textBoxName.Text.Trim();
            if (string.IsNullOrEmpty(textBoxName.Text))
            {
                MessageBox.Show("名称を入力してください。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

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