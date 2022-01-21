
namespace MarketingPDFMerge
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbPresentation = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.cmbPresenter1 = new System.Windows.Forms.ComboBox();
            this.cmbPresenter2 = new System.Windows.Forms.ComboBox();
            this.cmbDay = new System.Windows.Forms.ComboBox();
            this.cmbMonth = new System.Windows.Forms.ComboBox();
            this.cmbYear = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(124, 203);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(228, 23);
            this.button1.TabIndex = 8;
            this.button1.Text = "Create PDF";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(53, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Client Name";
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(125, 24);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(228, 20);
            this.txtClientName.TabIndex = 1;
            this.txtClientName.Text = "Client Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(56, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Presenter 2";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(56, 64);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(61, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Presenter 1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(87, 170);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(30, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Date";
            // 
            // cmbPresentation
            // 
            this.cmbPresentation.FormattingEnabled = true;
            this.cmbPresentation.Location = new System.Drawing.Point(124, 134);
            this.cmbPresentation.Name = "cmbPresentation";
            this.cmbPresentation.Size = new System.Drawing.Size(229, 21);
            this.cmbPresentation.TabIndex = 4;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(51, 137);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(66, 13);
            this.label7.TabIndex = 14;
            this.label7.Text = "Presentation";
            // 
            // cmbPresenter1
            // 
            this.cmbPresenter1.FormattingEnabled = true;
            this.cmbPresenter1.Location = new System.Drawing.Point(125, 61);
            this.cmbPresenter1.Name = "cmbPresenter1";
            this.cmbPresenter1.Size = new System.Drawing.Size(228, 21);
            this.cmbPresenter1.TabIndex = 2;
            // 
            // cmbPresenter2
            // 
            this.cmbPresenter2.FormattingEnabled = true;
            this.cmbPresenter2.Location = new System.Drawing.Point(125, 95);
            this.cmbPresenter2.Name = "cmbPresenter2";
            this.cmbPresenter2.Size = new System.Drawing.Size(228, 21);
            this.cmbPresenter2.TabIndex = 3;
            // 
            // cmbDay
            // 
            this.cmbDay.DropDownHeight = 130;
            this.cmbDay.FormattingEnabled = true;
            this.cmbDay.IntegralHeight = false;
            this.cmbDay.Location = new System.Drawing.Point(220, 167);
            this.cmbDay.Name = "cmbDay";
            this.cmbDay.Size = new System.Drawing.Size(45, 21);
            this.cmbDay.TabIndex = 6;
            // 
            // cmbMonth
            // 
            this.cmbMonth.FormattingEnabled = true;
            this.cmbMonth.Location = new System.Drawing.Point(125, 167);
            this.cmbMonth.Name = "cmbMonth";
            this.cmbMonth.Size = new System.Drawing.Size(89, 21);
            this.cmbMonth.TabIndex = 5;
            // 
            // cmbYear
            // 
            this.cmbYear.DropDownHeight = 130;
            this.cmbYear.FormattingEnabled = true;
            this.cmbYear.IntegralHeight = false;
            this.cmbYear.Location = new System.Drawing.Point(271, 167);
            this.cmbYear.Name = "cmbYear";
            this.cmbYear.Size = new System.Drawing.Size(81, 21);
            this.cmbYear.TabIndex = 7;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 241);
            this.Controls.Add(this.cmbYear);
            this.Controls.Add(this.cmbMonth);
            this.Controls.Add(this.cmbDay);
            this.Controls.Add(this.cmbPresenter2);
            this.Controls.Add(this.cmbPresenter1);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.cmbPresentation);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtClientName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Marketing Presentation";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbPresentation;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox cmbPresenter1;
        private System.Windows.Forms.ComboBox cmbPresenter2;
        private System.Windows.Forms.ComboBox cmbDay;
        private System.Windows.Forms.ComboBox cmbMonth;
        private System.Windows.Forms.ComboBox cmbYear;
    }
}

