
namespace OutlookAddIn_InsertLocalFileLinks
{
    partial class InsertForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listBoxInsertLocalFileLinks = new System.Windows.Forms.ListBox();
            this.checkBoxIsLinkOnlyDir = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // listBoxInsertLocalFileLinks
            // 
            this.listBoxInsertLocalFileLinks.AllowDrop = true;
            this.listBoxInsertLocalFileLinks.FormattingEnabled = true;
            this.listBoxInsertLocalFileLinks.ItemHeight = 12;
            this.listBoxInsertLocalFileLinks.Location = new System.Drawing.Point(12, 38);
            this.listBoxInsertLocalFileLinks.Name = "listBoxInsertLocalFileLinks";
            this.listBoxInsertLocalFileLinks.Size = new System.Drawing.Size(360, 160);
            this.listBoxInsertLocalFileLinks.TabIndex = 0;
            this.listBoxInsertLocalFileLinks.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBoxInsertLocalFileLinks_DragDrop);
            this.listBoxInsertLocalFileLinks.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBoxInsertLocalFileLinks_DragEnter);
            // 
            // checkBoxIsLinkOnlyDir
            // 
            this.checkBoxIsLinkOnlyDir.AutoSize = true;
            this.checkBoxIsLinkOnlyDir.Location = new System.Drawing.Point(13, 13);
            this.checkBoxIsLinkOnlyDir.Name = "checkBoxIsLinkOnlyDir";
            this.checkBoxIsLinkOnlyDir.Size = new System.Drawing.Size(130, 16);
            this.checkBoxIsLinkOnlyDir.TabIndex = 1;
            this.checkBoxIsLinkOnlyDir.Text = "フォルダだけリンクにする";
            this.checkBoxIsLinkOnlyDir.UseVisualStyleBackColor = true;
            this.checkBoxIsLinkOnlyDir.CheckedChanged += new System.EventHandler(this.checkBoxIsLinkOnlyDir_CheckedChanged);
            // 
            // InsertForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(384, 211);
            this.Controls.Add(this.checkBoxIsLinkOnlyDir);
            this.Controls.Add(this.listBoxInsertLocalFileLinks);
            this.Name = "InsertForm";
            this.Text = "挿入画面";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxInsertLocalFileLinks;
        private System.Windows.Forms.CheckBox checkBoxIsLinkOnlyDir;
    }
}