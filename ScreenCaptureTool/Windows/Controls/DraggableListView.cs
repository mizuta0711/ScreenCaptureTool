using System;
using System.Collections.ObjectModel;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows;
using System.Windows.Controls;

namespace ScreenCaptureTool.Windows.Controls
{
    /// <summary>
    /// ドラッグアンドドロップ可能なListView
    /// </summary>
    /// <typeparam name="T">データ型</typeparam>
    public class DraggableListView<T> : ListView where T : class
    {
        #region InnerClasses

        /// <summary>
        /// ドロップ位置のインジケータを描画するAdorner
        /// </summary>
        public class DropIndicatorAdorner : Adorner
        {
            private Point _currentPosition; // マウス位置

            public DropIndicatorAdorner(UIElement adornedElement)
                : base(adornedElement)
            {
            }

            // マウス位置を更新するメソッド
            public void UpdatePosition(Point position)
            {
                _currentPosition = position;
                InvalidateVisual(); // 再描画をトリガー
            }

            protected override void OnRender(DrawingContext drawingContext)
            {
                var listView = (ListView)AdornedElement;
                // 現在のマウス位置に最も近いListViewItemを取得
                if (listView.InputHitTest(_currentPosition) is FrameworkElement hitTestResult)
                {
                    var listViewItem = FindAncestor<ListViewItem>(hitTestResult);

                    if (listViewItem == null)
                    {
                        return;
                    }

                    // ListViewItem全体の領域を取得
                    var itemBounds = new Rect(listViewItem.TranslatePoint(new Point(0, 0), listView),
                                              new Size(listView.ActualWidth, listViewItem.ActualHeight));

                    // 下線を描画
                    var pen = new Pen(Brushes.Blue, 2);
                    drawingContext.DrawLine(pen, itemBounds.BottomLeft, itemBounds.BottomRight);
                }
            }

            // ヘルパーメソッド: 指定された型の親要素を取得
            private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
            {
                while (current != null)
                {
                    if (current is T ancestor)
                    {
                        return ancestor;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
                return null;
            }
        }

        #endregion InnerClasses

        #region Fields

        /// <summary>
        /// ドラッグ開始時のマウス位置
        /// </summary>
        private Point _dragStartPoint;

        /// <summary>
        /// ドロップ位置のインジケータを表示するためのAdornerLayer
        /// </summary>
        private AdornerLayer? _adornerLayer;

        /// <summary>
        /// 現在表示中のAdorner
        /// </summary>
        private DropIndicatorAdorner? _currentAdorner;

        #endregion Fields

        #region Constructor

        public DraggableListView()
        {
            // ドラッグアンドドロップのイベントハンドラを設定
            PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
            MouseMove += OnMouseMove;
            DragOver += OnDragOver;
            Drop += OnDrop;
            DragLeave += OnDragLeave;

            // ドラッグアンドドロップを有効にする
            AllowDrop = true;
        }

        #endregion Constructor

        #region Methods(Events)

        /// <summary>
        /// ドラッグ開始時のマウス位置を記録
        /// </summary>
        private void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
        }

        /// <summary>
        /// ドラッグ開始の処理
        /// </summary>
        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            // マウスの移動量を取得
            var mousePos = e.GetPosition(null);
            var diff = _dragStartPoint - mousePos;

            // ドラッグ開始判定（一定距離マウスが移動した場合のみドラッグ開始）
            if (Math.Abs(diff.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(diff.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            // ドラッグ開始
            var listView = (ListView)sender;
            if (listView.SelectedItem is T draggedItem)
            {
                DragDrop.DoDragDrop(listView, draggedItem, DragDropEffects.Move);
            }
        }

        /// <summary>
        /// ドラッグオーバー時にアイテムを受け入れるか確認
        /// </summary>
        private void OnDragOver(object sender, DragEventArgs e)
        {
            // マウス位置からドロップ先のアイテムを取得
            var listView = (ListView)sender;
            var point = e.GetPosition(listView);
            var targetItem = listView.InputHitTest(point) as FrameworkElement;
            var target = targetItem?.DataContext as T;

            if (target == null)
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            e.Effects = DragDropEffects.Move;

            // ドロップ位置のインジケータを表示
            // Adornerの更新処理
            if (_currentAdorner == null)
            {
                // 新規Adorner作成
                _adornerLayer = AdornerLayer.GetAdornerLayer(listView);
                if (_adornerLayer != null)
                {
                    _currentAdorner = new DropIndicatorAdorner(listView);
                    _adornerLayer.Add(_currentAdorner);
                }
            }
            _currentAdorner?.UpdatePosition(point);

            // ドロップされたデータがTである場合のみ受け入れる
            if (e.Data.GetDataPresent(typeof(T)))
            {
                e.Effects = DragDropEffects.Move;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        /// <summary>
        /// ドロップ時の処理（リストの項目を入れ替える）
        /// </summary>
        private void OnDrop(object sender, DragEventArgs e)
        {
            // ドロップ処理後にインジケータを削除
            if (_adornerLayer != null && _currentAdorner != null)
            {
                _adornerLayer.Remove(_currentAdorner);
                _adornerLayer = null;
                _currentAdorner = null;
            }

            // ドロップされたデータがTでない場合は何もしない
            if (!(e.Data.GetData(typeof(T)) is T draggedItem))
                return;

            // マウス位置からドロップ先のアイテムを取得
            var listView = (ListView)sender;
            var point = e.GetPosition(listView);
            var targetItem = listView.InputHitTest(point) as FrameworkElement;
            var target = targetItem?.DataContext as T;

            // ドロップ先がない場合やドロップ先がドラッグ元と同じ場合は何もしない
            if (target == null || draggedItem == target)
            {
                return;
            }

            // ドラッグ元とドロップ先のインデックスを取得
            var items = (ObservableCollection<T>)listView.ItemsSource;

            // ドラッグ元とドロップ先のインデックスを取得
            int oldIndex = items.IndexOf(draggedItem);
            int newIndex = items.IndexOf(target);

            // ドラッグ元をドロップ先の位置に挿入
            if (oldIndex != newIndex)
            {
                items.Move(oldIndex, newIndex);
            }
        }

        /// <summary>
        /// ドラッグがリストビュー外に出た場合にAdornerを削除
        /// </summary>
        private void OnDragLeave(object sender, DragEventArgs e)
        {
            // ドラッグがリストビュー外に出た場合にAdornerを削除
            if (_adornerLayer != null && _currentAdorner != null)
            {
                _adornerLayer.Remove(_currentAdorner);
                _adornerLayer = null;
                _currentAdorner = null;
            }
        }

        #endregion Methods(Events)
    }
}