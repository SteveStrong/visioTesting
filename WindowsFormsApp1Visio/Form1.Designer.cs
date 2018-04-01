namespace WindowsFormsApp1Visio
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
            this.axDrawingControl1 = new AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl();
            this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
            this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
            ((System.ComponentModel.ISupportInitialize)(this.axDrawingControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
            this.SuspendLayout();
            // 
            // axDrawingControl1
            // 
            this.axDrawingControl1.Enabled = true;
            this.axDrawingControl1.Location = new System.Drawing.Point(12, 12);
            this.axDrawingControl1.Name = "axDrawingControl1";
            this.axDrawingControl1.Size = new System.Drawing.Size(755, 601);
            this.axDrawingControl1.TabIndex = 0;
            // 
            // ultraTextEditor1
            // 
            this.ultraTextEditor1.Location = new System.Drawing.Point(787, 216);
            this.ultraTextEditor1.Multiline = true;
            this.ultraTextEditor1.Name = "ultraTextEditor1";
            this.ultraTextEditor1.Size = new System.Drawing.Size(301, 371);
            this.ultraTextEditor1.TabIndex = 2;
            // 
            // ultraButton1
            // 
            this.ultraButton1.Location = new System.Drawing.Point(801, 66);
            this.ultraButton1.Name = "ultraButton1";
            this.ultraButton1.Size = new System.Drawing.Size(75, 23);
            this.ultraButton1.TabIndex = 3;
            this.ultraButton1.Text = "Setup Visio";
            this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1228, 637);
            this.Controls.Add(this.ultraButton1);
            this.Controls.Add(this.ultraTextEditor1);
            this.Controls.Add(this.axDrawingControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.axDrawingControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl axDrawingControl1;
        private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
        private Infragistics.Win.Misc.UltraButton ultraButton1;
    }
}

