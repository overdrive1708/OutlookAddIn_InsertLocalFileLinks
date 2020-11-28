using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn_InsertLocalFileLinks
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility のメンバー

        public string GetCustomUI(string ribbonID)
        {
            string ribbonXML = String.Empty;

            // メッセージ編集画面のみ、アドインを表示する。
            if (ribbonID == "Microsoft.Outlook.Mail.Compose")
            {
                return GetResourceText("OutlookAddIn_InsertLocalFileLinks.Ribbon.xml");
            }

            return ribbonXML;
        }

        #endregion

        #region リボンのコールバック
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// 挿入画面起動ボタン押下時の処理
        /// </summary>
        public void OnInsertWindowButton(Office.IRibbonControl control)
        {
            InsertForm form = new InsertForm();
            form.ShowDialog();
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
