﻿using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;

namespace ScreenCaptureTool.Models
{
    /// <summary>
    /// プロジェクトの設定を保持するクラス
    /// </summary>
    [Serializable]
    public class ProjectSettings
    {
        #region Enum

        public enum CaptureType
        {
            ScreenRect,     // 画面の矩形
            Window          // ウィンドウ
        }

        #endregion Enum

        #region Properties

        /// <summary>
        /// プロジェクトファイルのパス
        /// </summary>
        public string FilePath;

        /// <summary>
        /// キャプチャー範囲：X
        /// </summary>
        public int CaptureLeft { get; set; }

        /// <summary>
        /// キャプチャー範囲：Y
        /// </summary>
        public int CaptureTop { get; set; }

        /// <summary>
        /// キャプチャー範囲：幅
        /// </summary>
        public int CaptureWidth { get; set; }

        /// <summary>
        /// キャプチャー範囲：高さ
        /// </summary>
        public int CaptureHeight { get; set; }

        /// <summary>
        /// ウィンドウ位置：X
        /// </summary>
        public double WindowLeft { get; set; }

        /// <summary>
        /// ウィンドウ位置：Y
        /// </summary>
        public double WindowTop { get; set; }

        /// <summary>
        /// ウィンドウサイズ：幅
        /// </summary>
        public double WindowWidth { get; set; }

        /// <summary>
        /// ウィンドウサイズ：高さ
        /// </summary>
        public double WindowHeight { get; set; }

        /// <summary>
        /// キャプチャーするウィンドウタイトル
        /// </summary>
        public string CaptureWindowTitle { get; set; }

        /// <summary>
        /// チャプチャータイプ
        /// </summary>
        public CaptureType SelectedCaptureType { get; set; }

        /// <summary>
        /// サムネイル画像サイズ
        /// </summary>
        public int ThumbnailSize { get; set; }

        /// <summary>
        /// 保存ファイル名の一覧
        /// </summary>
        public ObservableCollection<string> SaveFileNames;

        /// <summary>
        /// 画像保存先フォルダ
        /// </summary>
        public string SaveFolderPath { get; set; }

        #endregion Properties

        #region Constructor

        /// <summary>
        /// デフォルトコンストラクタ
        /// </summary>
        public ProjectSettings()
        {
            WindowLeft = 0;
            WindowTop = 0;
            WindowWidth = 0;
            WindowHeight = 0;

            SelectedCaptureType = CaptureType.ScreenRect;
            CaptureLeft = 0;
            CaptureTop = 0;
            CaptureWidth = 1920;
            CaptureHeight = 1280;
            CaptureWindowTitle = "";

            ThumbnailSize = 200;
            SaveFolderPath = Environment.CurrentDirectory;
            SaveFileNames = new ObservableCollection<string>();

            FilePath = Path.Combine(Environment.CurrentDirectory, "ScreenCaptureTool.scp");
        }

        #endregion Constructor

        #region Methods

        #region Methods(Public)

        /// <summary>
        /// ファイルに保存する
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>true: 成功 / false: 失敗</returns>
        public bool Save(string filePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ProjectSettings));
                using (var fs = new FileStream(filePath, FileMode.Create))
                {
                    serializer.Serialize(fs, this);
                }
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        #endregion Methods(Public)

        #region Methods(Static)

        /// <summary>
        /// ファイルから設定を読み込む
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        /// <returns>インスタンス(失敗時はnull)</returns>
        public static ProjectSettings? Load(string filePath)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(ProjectSettings));
                using (var fs = new FileStream(filePath, FileMode.Open))
                {
                    var projectSetting = serializer.Deserialize(fs) as ProjectSettings;
                    if (projectSetting != null)
                    {
                        projectSetting.FilePath = filePath;
                        return projectSetting;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return null;
        }

        #endregion Methods(Static)

        #endregion Methods
    }
}