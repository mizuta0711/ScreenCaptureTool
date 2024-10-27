using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ScreenCaptureTool.Windows.Controls
{
    internal class FolderTreeView : TreeView
    {
        public FolderTreeView() : base()
        {
            //// ドライブのノードを作成
            //foreach (var drive in System.IO.DriveInfo.GetDrives())
            //{
            //    // ドライブのノードを作成
            //    DriveTreeNode driveNode = new DriveTreeNode(drive);
            //    Items.Add(driveNode);
            //}
            InitializeFolderTree();
        }

        /// <summary>
        /// フォルダツリーを初期化する
        /// </summary>
        private void InitializeFolderTree()
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady)
                {
                    var item = new TreeViewItem { Header = drive.Name, Tag = drive.Name };
                    item.Items.Add(null);  // ダミーアイテム
                    item.Expanded += Folder_Expanded;
                    Items.Add(item);
                }
            }
        }

        // フォルダを展開したときの処理
        private void Folder_Expanded(object sender, RoutedEventArgs e)
        {
            var item = (TreeViewItem)sender;
            if (item.Items.Count == 1 && item.Items[0] == null)  // ダミーアイテムの確認
            {
                item.Items.Clear();
                try
                {
                    var folderPath = item.Tag.ToString();
                    if (String.IsNullOrEmpty(folderPath))
                    {
                        return;
                    }
                    // フォルダが存在しない場合は何もしない
                    if (!Directory.Exists(folderPath))
                    {
                        return;
                    }

                    // フォルダ内のサブフォルダを追加
                    var directories = Directory.GetDirectories(folderPath);
                    foreach (var directory in directories)
                    {
                        var subItem = new TreeViewItem { Header = Path.GetFileName(directory), Tag = directory };
                        subItem.Items.Add(null);  // ダミーアイテム
                        subItem.Expanded += Folder_Expanded;
                        item.Items.Add(subItem);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    // アクセス権限がない場合は何もしない
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// フォルダを再探索してツリーを更新する
        /// </summary>
        internal void RefreshSelectedFolderTree()
        {
            var selectedItem = SelectedItem as TreeViewItem;
            if (selectedItem != null)
            {
                // 現在選択されたフォルダの子要素をクリア
                selectedItem.Items.Clear();

                // 再度フォルダを展開し、ツリーに反映
                try
                {
                    if (selectedItem.Tag.ToString() is string folderPath)
                    {
                        var directories = Directory.GetDirectories(folderPath);
                        foreach (var directory in directories)
                        {
                            var subItem = new TreeViewItem { Header = Path.GetFileName(directory), Tag = directory };
                            subItem.Items.Add(null);  // ダミーアイテム
                            subItem.Expanded += Folder_Expanded;
                            selectedItem.Items.Add(subItem);
                        }
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    // アクセス権限がない場合は何もしない
                    Console.WriteLine(ex.Message);
                }
            }
        }

        /// <summary>
        /// 指定されたフォルダをツリーで選択する
        /// </summary>
        /// <param name="selectFolderPath">選択するフォルダ</param>
        internal void SelectFolderInTree(string selectFolderPath)
        {
            string[] pathParts = selectFolderPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            TreeViewItem? currentItem = null;

            foreach (TreeViewItem driveItem in Items)
            {
                if (driveItem.Tag.ToString() == pathParts[0] + Path.DirectorySeparatorChar) // ドライブ名の一致を確認
                {
                    currentItem = driveItem;
                    currentItem.IsExpanded = true; // ドライブを展開
                    break;
                }
            }

            // ドライブが見つからない場合は終了
            if (currentItem == null)
            {
                return;
            }

            // ドライブ以下のフォルダを順次展開していく
            for (int i = 1; i < pathParts.Length; i++)
            {
                bool found = false;
                foreach (TreeViewItem subItem in currentItem.Items)
                {
                    if (subItem.Header.ToString() == pathParts[i])
                    {
                        currentItem = subItem;
                        currentItem.IsExpanded = true; // フォルダを展開
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    return; // 見つからなければ終了
                }
            }

            currentItem.IsSelected = true; // 最後のフォルダを選択
        }
    }
}