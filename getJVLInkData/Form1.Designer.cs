namespace getJVLInkData
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.AxJVLink1 = new AxJVDTLabLib.AxJVLink();
            this.mnuConfig = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuConfJV = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.tmrDownload = new System.Windows.Forms.Timer(this.components);
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.btnGetJVData = new System.Windows.Forms.Button();
            this.prgJVRead = new System.Windows.Forms.ProgressBar();
            this.prgDownload = new System.Windows.Forms.ProgressBar();
            this.rtbData = new System.Windows.Forms.RichTextBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AxJVLink1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuConfig});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(528, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // AxJVLink1
            // 
            this.AxJVLink1.Enabled = true;
            this.AxJVLink1.Location = new System.Drawing.Point(460, 31);
            this.AxJVLink1.Name = "AxJVLink1";
            this.AxJVLink1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("AxJVLink1.OcxState")));
            this.AxJVLink1.Size = new System.Drawing.Size(60, 60);
            this.AxJVLink1.TabIndex = 1;
            // 
            // mnuConfig
            // 
            this.mnuConfig.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuConfJV});
            this.mnuConfig.Name = "mnuConfig";
            this.mnuConfig.Size = new System.Drawing.Size(58, 20);
            this.mnuConfig.Text = "設定(&C)";
            // 
            // mnuConfJV
            // 
            this.mnuConfJV.Name = "mnuConfJV";
            this.mnuConfJV.Size = new System.Drawing.Size(180, 22);
            this.mnuConfJV.Text = "JLinkの設定(&J)...";
            this.mnuConfJV.Click += new System.EventHandler(this.mnuConfJV_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(7, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(188, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "1.保存するフォルダを選択してください。";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(15, 56);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(45, 25);
            this.button1.TabIndex = 3;
            this.button1.Text = "選択";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.textBox1.Location = new System.Drawing.Point(60, 58);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(456, 23);
            this.textBox1.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(7, 94);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(238, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "2.調教データを取得する日付を選択してください。";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(16, 112);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 19);
            this.dateTimePicker1.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(7, 143);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(177, 15);
            this.label3.TabIndex = 2;
            this.label3.Text = "3.データ取得ボタンを押してください。";
            // 
            // btnGetJVData
            // 
            this.btnGetJVData.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGetJVData.Location = new System.Drawing.Point(16, 161);
            this.btnGetJVData.Name = "btnGetJVData";
            this.btnGetJVData.Size = new System.Drawing.Size(75, 44);
            this.btnGetJVData.TabIndex = 3;
            this.btnGetJVData.Text = "データ取得";
            this.btnGetJVData.UseVisualStyleBackColor = true;
            // 
            // prgJVRead
            // 
            this.prgJVRead.Location = new System.Drawing.Point(97, 184);
            this.prgJVRead.Name = "prgJVRead";
            this.prgJVRead.Size = new System.Drawing.Size(419, 21);
            this.prgJVRead.TabIndex = 6;
            // 
            // prgDownload
            // 
            this.prgDownload.Location = new System.Drawing.Point(97, 161);
            this.prgDownload.Name = "prgDownload";
            this.prgDownload.Size = new System.Drawing.Size(419, 21);
            this.prgDownload.TabIndex = 6;
            // 
            // rtbData
            // 
            this.rtbData.Font = new System.Drawing.Font("Meiryo UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.rtbData.Location = new System.Drawing.Point(16, 211);
            this.rtbData.Name = "rtbData";
            this.rtbData.Size = new System.Drawing.Size(500, 140);
            this.rtbData.TabIndex = 7;
            this.rtbData.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 361);
            this.Controls.Add(this.rtbData);
            this.Controls.Add(this.prgDownload);
            this.Controls.Add(this.prgJVRead);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnGetJVData);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.AxJVLink1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.AxJVLink1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private AxJVDTLabLib.AxJVLink AxJVLink1;
        private System.Windows.Forms.ToolStripMenuItem mnuConfig;
        private System.Windows.Forms.ToolStripMenuItem mnuConfJV;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Timer tmrDownload;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnGetJVData;
        private System.Windows.Forms.ProgressBar prgJVRead;
        private System.Windows.Forms.ProgressBar prgDownload;
        private System.Windows.Forms.RichTextBox rtbData;
    }
}

