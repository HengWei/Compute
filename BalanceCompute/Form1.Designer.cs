﻿namespace BalanceCompute
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
            label1 = new Label();
            openFileDialog1 = new OpenFileDialog();
            textBox1 = new TextBox();
            button1 = new Button();
            textBox3 = new TextBox();
            button3 = new Button();
            button2 = new Button();
            textBox2 = new TextBox();
            label2 = new Label();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(25, 31);
            label1.Name = "label1";
            label1.Size = new Size(67, 15);
            label1.TabIndex = 0;
            label1.Text = "匯入明細檔";
            label1.Click += label1_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(98, 31);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(358, 23);
            textBox1.TabIndex = 3;
            // 
            // button1
            // 
            button1.Location = new Point(480, 31);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 5;
            button1.Text = "匯入";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(33, 140);
            textBox3.Multiline = true;
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(423, 258);
            textBox3.TabIndex = 7;
            // 
            // button3
            // 
            button3.Location = new Point(480, 140);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 8;
            button3.Text = "開始處理";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button2
            // 
            button2.Location = new Point(480, 87);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 11;
            button2.Text = "匯入";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(98, 87);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(358, 23);
            textBox2.TabIndex = 10;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(25, 87);
            label2.Name = "label2";
            label2.Size = new Size(55, 15);
            label2.TabIndex = 9;
            label2.Text = "銀行交易";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(589, 423);
            Controls.Add(button2);
            Controls.Add(textBox2);
            Controls.Add(label2);
            Controls.Add(button3);
            Controls.Add(textBox3);
            Controls.Add(button1);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Name = "Form1";
            Text = "總表產生工具 v 1.1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private OpenFileDialog openFileDialog1;
        private TextBox textBox1;
        private Button button1;
        private TextBox textBox3;
        private Button button3;
        private Button button2;
        private TextBox textBox2;
        private Label label2;
    }
}