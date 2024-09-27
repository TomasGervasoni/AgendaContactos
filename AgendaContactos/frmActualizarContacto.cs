using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AgendaContactos
{
    public partial class frmActualizarContacto : Form
    {
        public frmActualizarContacto()
        {
            InitializeComponent();
        }

        // Método para cargar los datos del contacto en los campos del formulario
        public void CargarDatosContacto(string nombre, string apellido, string telefono, string correo, string categoria)
        {
            txtNombre.Text = nombre;
            txtApellido.Text = apellido;
            txtTelefono.Text = telefono;
            txtCorreo.Text = correo;
            cmbCategoria.SelectedItem = categoria;
        }

        // Evento del botón btnOK para guardar los cambios
        private void btnOK_Click(object sender, EventArgs e)
        {
            // Obtener los datos modificados
            string nombreModificado = txtNombre.Text;
            string apellidoModificado = txtApellido.Text;
            string telefonoModificado = txtTelefono.Text;
            string correoModificado = txtCorreo.Text;
            string categoriaModificada = cmbCategoria.SelectedItem.ToString();

            // Validar que no estén vacíos
            if (string.IsNullOrWhiteSpace(nombreModificado) ||
                string.IsNullOrWhiteSpace(apellidoModificado) ||
                string.IsNullOrWhiteSpace(telefonoModificado) ||
                string.IsNullOrWhiteSpace(correoModificado) ||
                string.IsNullOrWhiteSpace(categoriaModificada))
            {
                MessageBox.Show("Por favor, complete todos los campos.");
                return;
            }

            // Actualizar en la base de datos (requiere conexión a la base de datos)
            string query = "UPDATE Contactos SET Nombre = ?, Apellido = ?, Telefono = ?, Correo = ?, Categoria = ? WHERE Nombre = ? AND Apellido = ?";

            // Ruta a la base de datos Access
            string databasePath = "BaseDatos\\Contactos.accdb";
            ConexionBD conexionBD = new ConexionBD(databasePath);

            try
            {
                conexionBD.Abrir();

                // Crear el comando SQL
                using (OleDbCommand command = new OleDbCommand(query, conexionBD.ObtenerConexion()))
                {
                    // Añadir parámetros al comando
                    command.Parameters.AddWithValue("?", nombreModificado);
                    command.Parameters.AddWithValue("?", apellidoModificado);
                    command.Parameters.AddWithValue("?", telefonoModificado);
                    command.Parameters.AddWithValue("?", correoModificado);
                    command.Parameters.AddWithValue("?", categoriaModificada);
                    command.Parameters.AddWithValue("?", txtNombre.Text); // Claves originales
                    command.Parameters.AddWithValue("?", txtApellido.Text);

                    // Ejecutar el comando
                    command.ExecuteNonQuery();
                }

                // Mensaje de éxito
                MessageBox.Show("Datos actualizados correctamente.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al actualizar los datos: " + ex.Message);
            }
            finally
            {
                conexionBD.Cerrar();
            }

            // Cerrar el formulario después de actualizar
            this.Close();
        }

        private void frmActualizarContacto_Load(object sender, EventArgs e)
        {

        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
                 Close();
        }
    }
}
