namespace login
{
    public partial class Form1 : Form
    {
        Operaciones Operaciones = new Operaciones();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void contrasena_TextChanged(object sender, EventArgs e)
        {

        }

        private void lbUsuario_Click(object sender, EventArgs e)
        {

        }


        private void btIngresar_Click(object sender, EventArgs e)
        {
            string usuario = Nombre_usuario.Text;
            string contraseņa = contrasena.Text;
            if (usuario == "admin" && contraseņa == "1234")
            {
                MessageBox.Show("Contraseņa correcta");
                Operaciones.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Contraseņa incorrecta");
            }

        }
    }
}