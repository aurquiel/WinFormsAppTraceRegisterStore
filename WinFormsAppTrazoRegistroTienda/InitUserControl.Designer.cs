namespace WinFormsAppTrazoRegistroTienda
{
    partial class InitUserControl
    {
        /// <summary> 
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de componentes

        /// <summary> 
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.pictureBoxInit = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInit)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBoxInit
            // 
            this.pictureBoxInit.BackColor = System.Drawing.Color.Gainsboro;
            this.pictureBoxInit.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBoxInit.Image = global::WinFormsAppTrazoRegistroTienda.Properties.Resources.operations1;
            this.pictureBoxInit.Location = new System.Drawing.Point(0, 0);
            this.pictureBoxInit.Name = "pictureBoxInit";
            this.pictureBoxInit.Size = new System.Drawing.Size(1092, 648);
            this.pictureBoxInit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBoxInit.TabIndex = 1;
            this.pictureBoxInit.TabStop = false;
            // 
            // InitUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pictureBoxInit);
            this.Name = "InitUserControl";
            this.Size = new System.Drawing.Size(1092, 648);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxInit)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private PictureBox pictureBoxInit;
    }
}
