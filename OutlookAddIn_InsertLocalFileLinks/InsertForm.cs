using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace OutlookAddIn_InsertLocalFileLinks
{
    public partial class InsertForm : Form
    {
        public InsertForm()
        {
            InitializeComponent();

            // リストボックスの表示を設定
            listBoxInsertLocalFileLinks.Items.Add("ローカルファイルリンクを挿入したいフォルダ/ファイルを、");
            listBoxInsertLocalFileLinks.Items.Add("ここにドロップしてください。");

            // 設定を読み込む
            Properties.Settings.Default.Reload();

            // 設定値の表示反映
            checkBoxIsLinkOnlyDir.Checked = Properties.Settings.Default.IsLinkOnlyDir;
        }

        /// <summary>
        /// 挿入用リストボックスドラッグ時の処理
        /// </summary>
        private void listBoxInsertLocalFileLinks_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, true) == true)
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 挿入用リストボックスドロップ時の処理
        /// </summary>
        private void listBoxInsertLocalFileLinks_DragDrop(object sender, DragEventArgs e)
        {
            string strCurrentDirLink = "";
            string strInsertLink = "";

            Outlook.Application application = Globals.ThisAddIn.Application;
            Outlook.Inspector inspector = application.ActiveInspector();
            Outlook.MailItem mailItem = (Outlook.MailItem)inspector.CurrentItem;
            Word.Document doc = mailItem.GetInspector.WordEditor;

            if (e.Data.GetData(DataFormats.FileDrop) is string[] dropitems)
            {
                if (Properties.Settings.Default.IsLinkOnlyDir == true)
                {
                    // フォルダだけリンクにする場合

                    // ドラッグ&ドロップしたときのカレントディレクトリを取得
                    if (File.Exists(dropitems[0]) == true)
                    {
                        // ドロップしたアイテム(1つ目)がファイルの場合
                        strCurrentDirLink = Path.GetDirectoryName(dropitems[0]);
                    }
                    else if (Directory.Exists(dropitems[0]) == true)
                    {
                        // ドロップしたアイテム(1つ目)がフォルダの場合
                        DirectoryInfo di = new DirectoryInfo(dropitems[0]);
                        DirectoryInfo diParent = di.Parent;
                        strCurrentDirLink = diParent.FullName;
                    }
                    // リンクの生成
                    strCurrentDirLink = "<\"file://" + strCurrentDirLink + "\">";

                    // リンクの挿入
                    doc.Application.Selection.TypeText(strCurrentDirLink + "\r\n");

                    // フォルダ名/ファイル名の処理
                    foreach (string dropitem in dropitems)
                    {
                        // フォルダ名/ファイル名の生成
                        strInsertLink = "・" + Path.GetFileName(dropitem);

                        // フォルダ名/ファイル名の生成
                        doc.Application.Selection.TypeText(strInsertLink + "\r\n");
                    }
                }
                else
                {
                    // ファイルもリンクにする場合

                    foreach (string dropitem in dropitems)
                    {
                        // リンクの生成
                        strInsertLink = "<\"file://" + dropitem + "\">";

                        // リンクの挿入
                        doc.Application.Selection.TypeText(strInsertLink + "\r\n");
                    }
                }
            }

            // 画面を閉じる
            this.Close();
        }

        /// <summary>
        /// 設定用チェックボックス内容変更時の処理
        /// </summary>
        private void checkBoxIsLinkOnlyDir_CheckedChanged(object sender, EventArgs e)
        {
            // 設定を保存する
            Properties.Settings.Default.IsLinkOnlyDir = checkBoxIsLinkOnlyDir.Checked;
            Properties.Settings.Default.Save();
        }
    }
}
