using ScreenCaptureTool.Utilities;

using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Media3D;

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
                UpdateSelectionRectangle(startPoint, e.GetPosition(this));
            }
        }

        /// <summary>
        /// 選択中矩形を更新
        /// </summary>
        /// <param name="start">ドラッグ開始位置</param>
        /// <param name="end">ドラッグ終了位置</param>
        /// <returns>スクリーン座標での矩形</returns>
        private Rect UpdateSelectionRectangle(Point start, Point end)
        {
            var minPos = new Point(Math.Min(start.X, end.X), Math.Min(start.Y, end.Y));
            var maxPos = new Point(Math.Max(start.X, end.X), Math.Max(start.Y, end.Y));

            // 選択範囲矩形の位置とサイズを更新
            Canvas.SetLeft(SelectionRectangle, minPos.X);
            Canvas.SetTop(SelectionRectangle, minPos.Y);
            SelectionRectangle.Width = maxPos.X - minPos.X + 1;
            SelectionRectangle.Height = maxPos.Y - minPos.Y + 1;

            // スクリーン座標に変換
            return new Rect(this.PointToScreen(minPos), this.PointToScreen(maxPos));
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
                SelectedRect = UpdateSelectionRectangle(startPoint, e.GetPosition(this));

                // ダイアログを閉じて結果を返す
                DialogResult = true;
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