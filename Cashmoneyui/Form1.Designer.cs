
namespace Cashmoneyui
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonLoad = new System.Windows.Forms.Button();
            this.textBoxLoadTime = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBoxWriteTime = new System.Windows.Forms.TextBox();
            this.buttonWrite = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonLoad
            // 
            this.buttonLoad.Location = new System.Drawing.Point(95, 69);
            this.buttonLoad.Name = "buttonLoad";
            this.buttonLoad.Size = new System.Drawing.Size(94, 29);
            this.buttonLoad.TabIndex = 0;
            this.buttonLoad.Text = "Load File";
            this.buttonLoad.UseVisualStyleBackColor = true;
            this.buttonLoad.Click += new System.EventHandler(this.buttonLoad_Click);
            // 
            // textBoxLoadTime
            // 
            this.textBoxLoadTime.Location = new System.Drawing.Point(79, 104);
            this.textBoxLoadTime.Name = "textBoxLoadTime";
            this.textBoxLoadTime.Size = new System.Drawing.Size(125, 27);
            this.textBoxLoadTime.TabIndex = 1;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBoxWriteTime
            // 
            this.textBoxWriteTime.Location = new System.Drawing.Point(244, 104);
            this.textBoxWriteTime.Name = "textBoxWriteTime";
            this.textBoxWriteTime.Size = new System.Drawing.Size(125, 27);
            this.textBoxWriteTime.TabIndex = 2;
            // 
            // buttonWrite
            // 
            this.buttonWrite.Location = new System.Drawing.Point(259, 69);
            this.buttonWrite.Name = "buttonWrite";
            this.buttonWrite.Size = new System.Drawing.Size(94, 29);
            this.buttonWrite.TabIndex = 3;
            this.buttonWrite.Text = "Match";
            this.buttonWrite.UseVisualStyleBackColor = true;
            this.buttonWrite.Click += new System.EventHandler(this.buttonWrite_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(473, 286);
            this.Controls.Add(this.buttonWrite);
            this.Controls.Add(this.textBoxWriteTime);
            this.Controls.Add(this.textBoxLoadTime);
            this.Controls.Add(this.buttonLoad);
            this.Name = "Form1";
            this.Text = "FundIndexMatch";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonLoad;
        private System.Windows.Forms.TextBox textBoxLoadTime;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBoxWriteTime;
        private System.Windows.Forms.Button buttonWrite;
    }
}

