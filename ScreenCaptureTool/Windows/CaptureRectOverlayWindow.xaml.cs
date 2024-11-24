using ScreenCaptureTool.Models;
using ScreenCaptureTool.Models.CaptureItem;
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
        /// 位置固定フラグ
        /// </summary>
        private bool isFixedPos = false;

        /// <summary>
        /// サイズ固定フラグ
        /// </summary>
        private bool isFixedSize = false;

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

        #region Methods(Public)

        /// <summary>
        /// キャプチャーアイテムを設定
        /// </summary>
        /// <param name="captureItem">キャプチャーアイテム</param>
        public void SetCaptureItem(ManualScreenRectCaptureItem captureItem)
        {
            var location = new Point(0, 0);
            var size = new Size(0, 0);

            // 位置固定の場合はウィンドウ位置を設定
            isFixedPos = captureItem.LocationFixed;
            if (isFixedPos)
            {
                // 位置固定の場合はウィンドウ位置を設定
                location = new Point(captureItem.TargetRect.Left, captureItem.TargetRect.Top);
            }

            // サイズ固定の場合はウィンドウサイズを設定
            isFixedSize = captureItem.SizeFixed;
            if (isFixedSize)
            {
                // サイズ固定の場合はウィンドウサイズを設定
                size = new Size(captureItem.TargetRect.Width, captureItem.TargetRect.Height);
            }

            // 選択範囲を設定
            SelectedRect = new Rect(location, size);
        }

        #endregion Methods(Public)

        #region Methods(Event)

        /// <summary>
        /// 選択中矩形を更新
        /// </summary>
        /// <param name="start">ドラッグ開始位置</param>
        /// <param name="current">現在位置</param>
        /// <returns>スクリーン座標での矩形</returns>
        private Rect UpdateSelectionRectangle(Point start, Point current)
        {
            // 位置固定の場合
            if (isFixedPos)
            {
                // 開始位置を強制的に変更して固定する
                // スクリーン座標で保持されているので、ウィンドウ座標に変換
                start = PointFromScreen(SelectedRect.Location);
            }

            // サイズ固定の場合
            if (isFixedSize)
            {
                // 現在位置を開始位置として、サイズを固定する
                start = current;
                // スクリーン座標で保持されているので、ウィンドウ座標に変換
                var size = PointFromScreen(new Point(SelectedRect.Width - 1.0, SelectedRect.Height - 1.0));
                current = new Point(start.X + size.X, start.Y + size.Y);
            }

            // 選択範囲の矩形を計算
            var minPos = new Point(Math.Min(start.X, current.X), Math.Min(start.Y, current.Y));
            var maxPos = new Point(Math.Max(start.X, current.X), Math.Max(start.Y, current.Y));

            // 選択範囲矩形の位置とサイズを更新
            Canvas.SetLeft(SelectionRectangle, minPos.X);
            Canvas.SetTop(SelectionRectangle, minPos.Y);
            SelectionRectangle.Width = maxPos.X - minPos.X + 1;
            SelectionRectangle.Height = maxPos.Y - minPos.Y + 1;

            // スクリーン座標に変換
            return new Rect(PointToScreen(minPos), PointToScreen(maxPos));
        }

        /// <summary>
        /// 左クリックでドラッグ開始
        /// </summary>
        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // ドラッグ開始時の座標を記録
            startPoint = e.GetPosition(this);
            SelectionRectangle.Visibility = Visibility.Visible;
            UpdateSelectionRectangle(startPoint, startPoint);
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

        #endregion Methods(Event)

        #endregion Methods
    }
}