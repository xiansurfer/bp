namespace ArmpsCard_dll
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonArmpsCardIO = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // buttonArmpsCardIO
            // 
            this.buttonArmpsCardIO.Location = new System.Drawing.Point(24, 12);
            this.buttonArmpsCardIO.Name = "buttonArmpsCardIO";
            this.buttonArmpsCardIO.Size = new System.Drawing.Size(193, 43);
            this.buttonArmpsCardIO.TabIndex = 0;
            this.buttonArmpsCardIO.Text = "电流数据导入处理并传出";
            this.buttonArmpsCardIO.UseVisualStyleBackColor = true;
            this.buttonArmpsCardIO.Click += new System.EventHandler(this.buttonArmpsCardIO_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(395, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "注：测试时需要在C盘建立路径C:\\BPSeriesDemoTest\\ArmpsData\\的文件夹\r\n";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(617, 127);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.buttonArmpsCardIO);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonArmpsCardIO;
        private System.Windows.Forms.Label label1;
    }
}

