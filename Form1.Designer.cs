namespace login
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
            Nombre_usuario = new TextBox();
            lbUsuario = new Label();
            lbContrasena = new Label();
            contrasena = new TextBox();
            btIngresar = new Button();
            SuspendLayout();
            // 
            // Nombre_usuario
            // 
            Nombre_usuario.Location = new Point(29, 31);
            Nombre_usuario.Name = "Nombre_usuario";
            Nombre_usuario.Size = new Size(217, 23);
            Nombre_usuario.TabIndex = 0;
            Nombre_usuario.TextChanged += textBox1_TextChanged;
            // 
            // lbUsuario
            // 
            lbUsuario.AutoSize = true;
            lbUsuario.Location = new Point(116, 13);
            lbUsuario.Name = "lbUsuario";
            lbUsuario.Size = new Size(47, 15);
            lbUsuario.TabIndex = 1;
            lbUsuario.Text = "Usuario";
            lbUsuario.Click += lbUsuario_Click;
            // 
            // lbContrasena
            // 
            lbContrasena.AutoSize = true;
            lbContrasena.Location = new Point(105, 93);
            lbContrasena.Name = "lbContrasena";
            lbContrasena.Size = new Size(67, 15);
            lbContrasena.TabIndex = 2;
            lbContrasena.Text = "Contraseña";
            lbContrasena.Click += label2_Click;
            // 
            // contrasena
            // 
            contrasena.Location = new Point(29, 133);
            contrasena.Name = "contrasena";
            contrasena.Size = new Size(217, 23);
            contrasena.TabIndex = 3;
            contrasena.TextChanged += contrasena_TextChanged;
            // 
            // btIngresar
            // 
            btIngresar.Location = new Point(105, 200);
            btIngresar.Name = "btIngresar";
            btIngresar.Size = new Size(75, 23);
            btIngresar.TabIndex = 4;
            btIngresar.Text = "Ingresar";
            btIngresar.UseVisualStyleBackColor = true;
            btIngresar.Click += btIngresar_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Info;
            ClientSize = new Size(284, 361);
            Controls.Add(btIngresar);
            Controls.Add(contrasena);
            Controls.Add(lbContrasena);
            Controls.Add(lbUsuario);
            Controls.Add(Nombre_usuario);
            Name = "Form1";
            Text = "Login";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox Nombre_usuario;
        private Label lbUsuario;
        private Label lbContrasena;
        private TextBox contrasena;
        private Button btIngresar;
    }
}