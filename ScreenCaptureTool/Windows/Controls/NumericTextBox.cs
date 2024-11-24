using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ScreenCaptureTool.Windows.Controls
{
    /// <summary>
    /// 数値入力のみを受け付けるTextBox
    /// </summary>
    public class NumericTextBox : TextBox
    {
        public NumericTextBox()
        {
            // フォーカス時にテキストを全選択
            GotFocus += SelectAllOnFocus;

            // マウスクリック時にフォーカスをセットし全選択
            PreviewMouseDown += SelectAllOnMouseDown;
        }

        private void SelectAllOnFocus(object sender, RoutedEventArgs e)
        {
            this.SelectAll();
        }

        private void SelectAllOnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!IsKeyboardFocusWithin) // 既にフォーカスがある場合は処理しない
            {
                e.Handled = true; // 既定のマウス動作を抑制
                Focus(); // フォーカスをセット
                SelectAll(); // 全選択
            }
        }
    
        protected override void OnPreviewTextInput(TextCompositionEventArgs e)
        {
            base.OnPreviewTextInput(e);

            // 数値にパースできない文字列は入力を受け付けない
            e.Handled = !IsTextNumeric(e.Text);
        }

        private bool IsTextNumeric(string text)
        {
            return int.TryParse(text, out _);
        }
    }
}
