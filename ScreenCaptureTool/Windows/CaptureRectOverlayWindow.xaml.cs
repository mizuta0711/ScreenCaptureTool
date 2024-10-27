using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ScreenCaptureTool.Windows
{
    /// <summary>
    /// キャプチャー範囲選択用オーバーレイウィンドウ
    /// </summary>
    public partial class CaptureRectOverlayWindow : Window
    {
        #region Fields

        /// <summary>
        /// ドラッグ開始座標
        /// </summary>
        private Point startPoint;

        /// <summary>
        /// ドラッグ中かどうか
        /// </summary>
        private bool isDragging = false;

        #endregion Fields

        #region Properties

        /// <summary>
        /// 選択された矩形
        /// </summary>
        public Rect SelectedRect { get; private set; }

        #endregion Properties

        #region Constructor

        public CaptureRectOverlayWindow()
        {
            InitializeComponent();
        }

        #endregion Constructor

        #region Methods

        #region EventHandlers

        /// <summary>
        /// 左クリックでドラッグ開始
        /// </summary>
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // ドラッグ開始時の座標を記録
            startPoint = e.GetPosition(this);
            SelectionRectangle.Visibility = Visibility.Visible;
            SelectionRectangle.Width = 0;
            SelectionRectangle.Height = 0;
            isDragging = true;
        }

        /// <summary>
        /// マウス移動中に矩形を更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                // 現在のマウス位置を取得し、矩形を更新
                var pos = e.GetPosition(this);
                var x = Math.Min(pos.X, startPoint.X);
                var y = Math.Min(pos.Y, startPoint.Y);
                var width = Math.Abs(pos.X - startPoint.X);
                var height = Math.Abs(pos.Y - startPoint.Y);

                // 矩形のサイズと位置を更新
                Canvas.SetLeft(SelectionRectangle, x);
                Canvas.SetTop(SelectionRectangle, y);
                SelectionRectangle.Width = width;
                SelectionRectangle.Height = height;
            }
        }

        /// <summary>
        /// 左クリックでドラッグ終了
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (isDragging)
            {
                isDragging = false;
                // 選択範囲を確定
                var pos = e.GetPosition(this);
                var x = Math.Min(pos.X, startPoint.X);
                var y = Math.Min(pos.Y, startPoint.Y);
                var width = Math.Abs(pos.X - startPoint.X);
                var height = Math.Abs(pos.Y - startPoint.Y);

                SelectedRect = new Rect(x, y, width, height);
                DialogResult = true;  // ダイアログを閉じて結果を返す
                this.Close();
            }
        }

        /// <summary>
        /// キーイベントでESCが押されたらキャンセル
        /// </summary>
        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                // キャンセル処理
                isDragging = false;
                SelectionRectangle.Visibility = Visibility.Collapsed;
                DialogResult = false;  // キャンセルと判断
                this.Close();
            }
        }

        #endregion EventHandlers

        #endregion Methods
    }
}