using System.Windows;

namespace ScreenCaptureTool.Windows
{
    /// <summary>
    /// CaptureSettingsWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class CaptureSettingsWindow : Window
    {
        public CaptureSettingsWindow()
        {
            InitializeComponent();
        }

        #region Methods(Private)

        #region Methods(Event)

        /// <summary>
        /// 撮影方法のラジオボタン：選択変更
        /// </summary>
        private void RecordingType_CheckedChanged(object sender, RoutedEventArgs e)
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
            }
        }

        /// <summary>
        /// 撮影位置固定チェックボックス：チェック状態変更
        /// </summary>
        private void FixedPositionCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            recordingPositionPanel.IsEnabled = checkBoxFixedPos.IsChecked ?? false;
        }

        /// <summary>
        /// 撮影サイズ固定チェックボックス：チェック状態変更
        /// </summary>
        private void FixedSizeCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            recordingSizePanel.IsEnabled = checkBoxFixedSize.IsChecked ?? false;
        }

        /// <summary>
        /// 縁カットチェックボックス：チェック状態変更
        /// </summary>
        private void TrimEdgesCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            trimEdgesPanel.IsEnabled = checkBoxTrimEdge.IsChecked ?? false;
        }

        /// <summary>
        /// リサイズチェックボックス：チェック状態変更
        /// </summary>
        private void ResizeCheckBox_Clicked(object sender, RoutedEventArgs e)
        {
            resizePanel.IsEnabled = checkBoxResize.IsChecked ?? false;
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