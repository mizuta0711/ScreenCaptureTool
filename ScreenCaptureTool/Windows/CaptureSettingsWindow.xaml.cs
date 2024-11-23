using ScreenCaptureTool.Models;
using ScreenCaptureTool.Models.CaptureItem;

using System.Windows;

namespace ScreenCaptureTool.Windows
{
    /// <summary>
    /// CaptureSettingsWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class CaptureSettingsWindow : Window
    {
        #region Properties

        private CaptureItemBase CaptureItem { get; set; }

        #endregion Properties

        #region Constructors

        public CaptureSettingsWindow()
        {
            InitializeComponent();
            CaptureItem = new ManualScreenRectCaptureItem(new System.Drawing.Point(0, 0));
            LoadRecordingSetting(new RecordingSetting("", CaptureItem));
        }

        #endregion Constructors

        #region Methods(Private)

        private void LoadRecordingSetting(RecordingSetting setting)
        {
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

            // UIコントロールの更新
            RefreshUIControls();
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
        /// 撮影方法のラジオボタン：選択変更
        /// </summary>
        private void RecordingType_CheckedChanged(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// 撮影位置固定チェックボックス：チェック状態変更
        /// </summary>
        private void FixedPositionCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// 撮影サイズ固定チェックボックス：チェック状態変更
        /// </summary>
        private void FixedSizeCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// 縁トリミングチェックボックス：チェック状態変更
        /// </summary>
        private void TrimEdgeCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// リサイズチェックボックス：チェック状態変更
        /// </summary>
        private void ResizeCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            RefreshUIControls();
        }

        /// <summary>
        /// OKボタン：押下イベント
        /// </summary>
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
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
    }
}