namespace login
{
    partial class Operaciones
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
            label3 = new Label();
            label4 = new Label();
            lblFecha = new Label();
            groupBox1 = new GroupBox();
            lblPrecio = new Label();
            label10 = new Label();
            label9 = new Label();
            label8 = new Label();
            label7 = new Label();
            label6 = new Label();
            tbCantidad = new TextBox();
            cboTipo = new ComboBox();
            cboproducto = new ComboBox();
            listView1 = new ListView();
            columnHeader1 = new ColumnHeader();
            columnHeader2 = new ColumnHeader();
            columnHeader3 = new ColumnHeader();
            columnHeader4 = new ColumnHeader();
            columnHeader5 = new ColumnHeader();
            columnHeader6 = new ColumnHeader();
            columnHeader7 = new ColumnHeader();
            btnSalir = new Button();
            btnCancelar = new Button();
            btnRegistrar = new Button();
            btnGuardar = new Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("HP Simplified", 21.7499962F, FontStyle.Bold, GraphicsUnit.Point);
            label3.Location = new Point(207, 9);
            label3.Name = "label3";
            label3.Size = new Size(378, 34);
            label3.TabIndex = 0;
            label3.Text = "Control de registro de ventas";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label4.Location = new Point(647, 39);
            label4.Name = "label4";
            label4.Size = new Size(53, 18);
            label4.TabIndex = 1;
            label4.Text = "Fecha: ";
            // 
            // lblFecha
            // 
            lblFecha.AutoSize = true;
            lblFecha.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            lblFecha.Location = new Point(727, 39);
            lblFecha.Name = "lblFecha";
            lblFecha.Size = new Size(61, 18);
            lblFecha.TabIndex = 2;
            lblFecha.Text = "lblFecha";
            lblFecha.Click += label5_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(lblPrecio);
            groupBox1.Controls.Add(label10);
            groupBox1.Controls.Add(label9);
            groupBox1.Controls.Add(label8);
            groupBox1.Controls.Add(label7);
            groupBox1.Controls.Add(label6);
            groupBox1.Controls.Add(tbCantidad);
            groupBox1.Controls.Add(cboTipo);
            groupBox1.Controls.Add(cboproducto);
            groupBox1.Location = new Point(12, 71);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(764, 159);
            groupBox1.TabIndex = 3;
            groupBox1.TabStop = false;
            groupBox1.Text = "Area de ventas";
            groupBox1.Enter += groupBox1_Enter;
            // 
            // lblPrecio
            // 
            lblPrecio.AutoSize = true;
            lblPrecio.Font = new Font("HP Simplified", 11.2499981F, FontStyle.Bold, GraphicsUnit.Point);
            lblPrecio.Location = new Point(589, 88);
            lblPrecio.Name = "lblPrecio";
            lblPrecio.Size = new Size(66, 19);
            lblPrecio.TabIndex = 12;
            lblPrecio.Text = "lblPrecio";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label10.Location = new Point(598, 70);
            label10.Name = "label10";
            label10.Size = new Size(0, 18);
            label10.TabIndex = 11;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label9.Location = new Point(551, 42);
            label9.Name = "label9";
            label9.Size = new Size(141, 18);
            label9.TabIndex = 10;
            label9.Text = "Precio del producto:";
            label9.Click += label9_Click;
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label8.Location = new Point(308, 31);
            label8.Name = "label8";
            label8.Size = new Size(72, 18);
            label8.TabIndex = 9;
            label8.Text = "Cantidad:";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label7.Location = new Point(83, 96);
            label7.Name = "label7";
            label7.Size = new Size(92, 18);
            label7.TabIndex = 8;
            label7.Text = "Tipo de pago";
            label7.Click += label7_Click;
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Constantia", 11.25F, FontStyle.Regular, GraphicsUnit.Point);
            label6.Location = new Point(51, 31);
            label6.Name = "label6";
            label6.Size = new Size(166, 18);
            label6.TabIndex = 8;
            label6.Text = "Seleccione un producto:";
            // 
            // tbCantidad
            // 
            tbCantidad.Location = new Point(288, 61);
            tbCantidad.Name = "tbCantidad";
            tbCantidad.Size = new Size(114, 23);
            tbCantidad.TabIndex = 2;
            tbCantidad.TextChanged += tbCantidad_TextChanged;
            // 
            // cboTipo
            // 
            cboTipo.FormattingEnabled = true;
            cboTipo.Items.AddRange(new object[] { "Visa / MasterCard", "Efectivo" });
            cboTipo.Location = new Point(69, 117);
            cboTipo.Name = "cboTipo";
            cboTipo.Size = new Size(121, 23);
            cboTipo.TabIndex = 1;
            cboTipo.SelectedIndexChanged += comboBox2_SelectedIndexChanged;
            // 
            // cboproducto
            // 
            cboproducto.FormattingEnabled = true;
            cboproducto.Items.AddRange(new object[] { "Taco", "burrito", "Hamburguesa", "Pirata", "Quesadilla", "Empalme", "Sincronizada", "Agua", "Jugo", "Refresco" });
            cboproducto.Location = new Point(69, 61);
            cboproducto.Name = "cboproducto";
            cboproducto.Size = new Size(121, 23);
            cboproducto.TabIndex = 0;
            cboproducto.SelectedIndexChanged += cboproducto_SelectedIndexChanged;
            // 
            // listView1
            // 
            listView1.Columns.AddRange(new ColumnHeader[] { columnHeader1, columnHeader2, columnHeader3, columnHeader4, columnHeader5, columnHeader6, columnHeader7 });
            listView1.GridLines = true;
            listView1.Location = new Point(25, 292);
            listView1.Name = "listView1";
            listView1.Size = new Size(751, 179);
            listView1.TabIndex = 4;
            listView1.UseCompatibleStateImageBehavior = false;
            listView1.View = View.Details;
            listView1.SelectedIndexChanged += listView1_SelectedIndexChanged_1;
            // 
            // columnHeader1
            // 
            columnHeader1.Text = "Producto";
            columnHeader1.Width = 107;
            // 
            // columnHeader2
            // 
            columnHeader2.Text = "Cantidad";
            columnHeader2.TextAlign = HorizontalAlignment.Center;
            columnHeader2.Width = 107;
            // 
            // columnHeader3
            // 
            columnHeader3.Text = "Precio";
            columnHeader3.TextAlign = HorizontalAlignment.Center;
            columnHeader3.Width = 107;
            // 
            // columnHeader4
            // 
            columnHeader4.Text = "Tipo de Pago";
            columnHeader4.TextAlign = HorizontalAlignment.Center;
            columnHeader4.Width = 107;
            // 
            // columnHeader5
            // 
            columnHeader5.Text = "Decuento";
            columnHeader5.TextAlign = HorizontalAlignment.Center;
            columnHeader5.Width = 107;
            // 
            // columnHeader6
            // 
            columnHeader6.Text = "Recargo";
            columnHeader6.TextAlign = HorizontalAlignment.Center;
            columnHeader6.Width = 107;
            // 
            // columnHeader7
            // 
            columnHeader7.Text = "Precio Total";
            columnHeader7.Width = 107;
            // 
            // btnSalir
            // 
            btnSalir.BackColor = Color.FromArgb(255, 128, 0);
            btnSalir.Location = new Point(696, 477);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(92, 37);
            btnSalir.TabIndex = 5;
            btnSalir.Text = "Salir";
            btnSalir.UseVisualStyleBackColor = false;
            btnSalir.Click += button1_Click;
            // 
            // btnCancelar
            // 
            btnCancelar.BackColor = Color.FromArgb(255, 128, 0);
            btnCancelar.Location = new Point(284, 245);
            btnCancelar.Name = "btnCancelar";
            btnCancelar.Size = new Size(82, 41);
            btnCancelar.TabIndex = 6;
            btnCancelar.Text = "Cancelar";
            btnCancelar.UseVisualStyleBackColor = false;
            btnCancelar.Click += btnCancelar_Click;
            // 
            // btnRegistrar
            // 
            btnRegistrar.BackColor = Color.FromArgb(255, 128, 0);
            btnRegistrar.Location = new Point(402, 245);
            btnRegistrar.Name = "btnRegistrar";
            btnRegistrar.Size = new Size(80, 41);
            btnRegistrar.TabIndex = 7;
            btnRegistrar.Text = "Registrar";
            btnRegistrar.UseVisualStyleBackColor = false;
            btnRegistrar.Click += btnRegistrar_Click;
            // 
            // btnGuardar
            // 
            btnGuardar.BackColor = Color.FromArgb(255, 128, 0);
            btnGuardar.Location = new Point(575, 477);
            btnGuardar.Name = "btnGuardar";
            btnGuardar.Size = new Size(92, 37);
            btnGuardar.TabIndex = 8;
            btnGuardar.Text = "Guardar";
            btnGuardar.UseVisualStyleBackColor = false;
            btnGuardar.Click += btnGuardar_Click;
            // 
            // Operaciones
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(810, 517);
            Controls.Add(btnGuardar);
            Controls.Add(btnRegistrar);
            Controls.Add(btnCancelar);
            Controls.Add(btnSalir);
            Controls.Add(listView1);
            Controls.Add(groupBox1);
            Controls.Add(lblFecha);
            Controls.Add(label4);
            Controls.Add(label3);
            Name = "Operaciones";
            Text = "Registro_ventas";
            Load += Operaciones_Load;
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        #endregion

        private Label lblPrecio;
        private Label label2;
        private TextBox tbPrimero;
        private TextBox tbSegundo;
        private TextBox tbResultado;
        private Button btCalcular;
        private ComboBox cBoperaciones;
        private ListView lV_historial;
        private Label label3;
        private Label label4;
        private Label lblFecha;
        private GroupBox groupBox1;
        private Label label10;
        private Label label9;
        private Label label8;
        private Label label7;
        private Label label6;
        private TextBox tbCantidad;
        private ComboBox cboTipo;
        private ComboBox cboproducto;
        private ListView listView1;
        private Button btnSalir;
        private Button btnCancelar;
        private Button btnRegistrar;
        private ColumnHeader columnHeader1;
        private ColumnHeader columnHeader2;
        private ColumnHeader columnHeader3;
        private ColumnHeader columnHeader4;
        private ColumnHeader columnHeader5;
        private ColumnHeader columnHeader6;
        private ColumnHeader columnHeader7;
        private Button btnGuardar;
    }
}