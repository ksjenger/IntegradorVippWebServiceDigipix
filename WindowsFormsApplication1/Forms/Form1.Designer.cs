using System.Windows.Forms;

namespace IntegradorWebService
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labelPath = new System.Windows.Forms.Label();
            this.labelProgresso = new System.Windows.Forms.Label();
            this.comboPerfil = new System.Windows.Forms.ComboBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnEnviar = new System.Windows.Forms.Button();
            this.btnSelecione = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            resources.ApplyResources(this.openFileDialog1, "openFileDialog1");
            // 
            // labelPath
            // 
            resources.ApplyResources(this.labelPath, "labelPath");
            this.labelPath.Name = "labelPath";
            // 
            // labelProgresso
            // 
            resources.ApplyResources(this.labelProgresso, "labelProgresso");
            this.labelProgresso.Name = "labelProgresso";
            // 
            // comboPerfil
            // 
            this.comboPerfil.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboPerfil.FormattingEnabled = true;
            resources.ApplyResources(this.comboPerfil, "comboPerfil");
            this.comboPerfil.Name = "comboPerfil";
            this.comboPerfil.SelectedIndexChanged += new System.EventHandler(this.ComboPerfil_SelectedIndexChanged);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox1.Image = global::IntegradorWebService.Properties.Resources.logo_visualset1;
            resources.ApplyResources(this.pictureBox1, "pictureBox1");
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.PictureBox1_Click);
            // 
            // btnEnviar
            // 
            this.btnEnviar.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnEnviar.Image = global::IntegradorWebService.Properties.Resources.Send;
            resources.ApplyResources(this.btnEnviar, "btnEnviar");
            this.btnEnviar.Name = "btnEnviar";
            this.btnEnviar.UseVisualStyleBackColor = true;
            this.btnEnviar.Click += new System.EventHandler(this.Button2_Click);
            // 
            // btnSelecione
            // 
            this.btnSelecione.Cursor = System.Windows.Forms.Cursors.Default;
            this.btnSelecione.Image = global::IntegradorWebService.Properties.Resources.Open;
            resources.ApplyResources(this.btnSelecione, "btnSelecione");
            this.btnSelecione.Name = "btnSelecione";
            this.btnSelecione.Click += new System.EventHandler(this.Button1_Click_1);
            // 
            // Form1
            // 
            resources.ApplyResources(this, "$this");
            this.AutoValidate = System.Windows.Forms.AutoValidate.Disable;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.btnSelecione);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.comboPerfil);
            this.Controls.Add(this.btnEnviar);
            this.Controls.Add(this.labelProgresso);
            this.Controls.Add(this.labelPath);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

  
        #endregion

        
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label labelPath;
        public System.Windows.Forms.Label labelProgresso;
        private System.Windows.Forms.Button btnEnviar;
        private System.Windows.Forms.ComboBox comboPerfil;
        private System.Windows.Forms.PictureBox pictureBox1;
        private Button btnSelecione;
    }
}

