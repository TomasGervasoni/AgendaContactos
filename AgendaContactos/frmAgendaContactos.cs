using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AgendaContactos
{
    public partial class frmAgendaContactos : Form
    {
        public frmAgendaContactos()
        {
            InitializeComponent();
            CargarDatos();
        }
        private void frmAgendaContactos_Load(object sender, EventArgs e)
        {
            // Crea una instancia de la clase Conexion
            ConexionBD conexion = new ConexionBD("BaseDatos\\Contactos.accdb");
            conexion.Abrir();
            // Reinicializar el ComboBox al primer elemento
            if (cmbCategoria.Items.Count > 0)
            {
                cmbCategoria.SelectedIndex = 0;
            }

        }
        private void CargarDatos()
        {
            // Ruta a la base de datos Access
            string databasePath = ("BaseDatos\\Contactos.accdb");

            // Crea una instancia de la clase Conexion
            ConexionBD conexionBD = new ConexionBD(databasePath);

            try
            {
                // Abre la conexion
                conexionBD.Abrir();
                // Crea un adaptador de datos para llenar un DataTable
                string query = "SELECT * FROM Contactos";
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, conexionBD.ObtenerConexion());
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                // Asigna el DataTable al DataGridView
                dgvContactos.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar los datos: " + ex.Message);
            }
        }
        private void Cargar_Click(object sender, EventArgs e)
        {
            // Ruta a la base de datos Access
            string databasePath = "BaseDatos\\Contactos.accdb";

            // Crea una instancia de la clase Conexion
            ConexionBD conexionBD = new ConexionBD(databasePath);

            // Obtén los valores de los TextBox
            string nombre = txtNombre.Text;
            string Apellido = txtApellido.Text;
            string Telefono = txtTelefono.Text;
            string Correo = txtCorreo.Text;
            string categoria = cmbCategoria.SelectedItem?.ToString(); // Usar el valor seleccionado

            // Validar los valores (podrías agregar más validaciones según sea necesario)
            if (string.IsNullOrWhiteSpace(nombre) ||
                    string.IsNullOrWhiteSpace(Apellido) ||
                    string.IsNullOrWhiteSpace(Telefono) ||
                    string.IsNullOrWhiteSpace(Correo) ||
                    string.IsNullOrWhiteSpace(categoria))
            {
                MessageBox.Show("Por favor, complete todos los campos.");
                return;
            }

            // Crear la consulta SQL para insertar datos
            string query = "INSERT INTO Contactos (Nombre, Apellido, Telefono, Correo, Categoria) VALUES (?, ?, ?, ?, ?)";

            try
            {
                // Abre la conexión
                conexionBD.Abrir();

                // Crear el comando SQL
                using (OleDbCommand command = new OleDbCommand(query, conexionBD.ObtenerConexion()))
                {
                    // Añadir parámetros al comando
                    command.Parameters.AddWithValue("?", nombre);
                    command.Parameters.AddWithValue("?", Apellido);
                    command.Parameters.AddWithValue("?", Telefono);
                    command.Parameters.AddWithValue("?", Correo);
                    command.Parameters.AddWithValue("?", categoria);

                    // Ejecutar el comando
                    command.ExecuteNonQuery();
                }

                // Mensaje de éxito
                MessageBox.Show("Datos cargados correctamente.");

                // Limpiar los TextBox
                txtNombre.Clear();
                txtTelefono.Clear();
                txtCorreo.Clear();
                txtApellido.Clear();

                // Reinicializar el ComboBox al primer elemento
                if (cmbCategoria.Items.Count > 0)
                {
                    cmbCategoria.SelectedIndex = 0;
                }

                // Recargar los datos en el DataGridView
                CargarDatos();
            }
            catch (Exception ex)
            {
                // Mostrar mensaje de error
                MessageBox.Show("Error al cargar los datos: " + ex.Message);
            }
            finally
            {
                // Asegurarse de cerrar la conexión
                conexionBD.Cerrar();
            }
        }



        private void btnBuscar_Click(object sender, EventArgs e)
        {
            // Crear instancia del formulario frmBuscarContactos
            frmBuscarContactos frmBuscar = new frmBuscarContactos();

            // Pasar los valores de búsqueda al formulario
            frmBuscar.FiltrarDatos(txtNombre.Text, txtTelefono.Text, txtCorreo.Text);

            // Oculta el formulario frmAgendaContactos
            this.Hide();

            // Muestra el formulario frmBuscarContactos y verifica si se ha cerrado
            if (frmBuscar.ShowDialog() == DialogResult.OK)
            {
                // Actualiza la grilla de contactos (CargarDatos)
                CargarDatos();
            }

            // Vuelve a mostrar el formulario frmAgendaContactos
            this.Show();
        }

        private string ExportarCSV()
        {
            // Ruta de la carpeta "exel" dentro del directorio del proyecto
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string folderPath = Path.Combine(projectDirectory, "exel");

            // Crear la carpeta "exel" si no existe
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Definir la ruta del archivo CSV
            string filePath = Path.Combine(folderPath, "Contactos.csv");

            try
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    // Escribir las cabeceras
                    sw.WriteLine("Nombre,Apellido,Telefono,Correo,Categoria");

                    // Obtener los datos de los contactos desde el DataGridView
                    foreach (DataGridViewRow row in dgvContactos.Rows)
                    {
                        if (row.Cells[0].Value != null) // Asegurarse de que la fila no esté vacía
                        {
                            string nombre = row.Cells["Nombre"].Value.ToString();
                            string apellido = row.Cells["Apellido"].Value.ToString();
                            string telefono = row.Cells["Telefono"].Value.ToString();
                            string correo = row.Cells["Correo"].Value.ToString();
                            string categoria = row.Cells["Categoria"].Value.ToString();

                            // Escribir la fila en el archivo CSV
                            sw.WriteLine($"{nombre},{apellido},{telefono},{correo},{categoria}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a CSV: " + ex.Message);
            }

            return filePath;  // Devolver la ruta del archivo generado
        }

        private string ExportarVCard()
        {
            // Ruta de la carpeta "exel" dentro del directorio del proyecto
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string folderPath = Path.Combine(projectDirectory, "exel");

            // Crear la carpeta "exel" si no existe
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Definir la ruta del archivo vCard
            string filePath = Path.Combine(folderPath, "Contactos.vcf");

            try
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    foreach (DataGridViewRow row in dgvContactos.Rows)
                    {
                        if (row.Cells[0].Value != null) // Asegurarse de que la fila no esté vacía
                        {
                            string nombre = row.Cells["Nombre"].Value.ToString();
                            string apellido = row.Cells["Apellido"].Value.ToString();
                            string telefono = row.Cells["Telefono"].Value.ToString();
                            string correo = row.Cells["Correo"].Value.ToString();

                            // Escribir el formato vCard
                            sw.WriteLine("BEGIN:VCARD");
                            sw.WriteLine("VERSION:3.0");
                            sw.WriteLine($"FN:{nombre} {apellido}");
                            sw.WriteLine($"TEL;TYPE=CELL:{telefono}");
                            sw.WriteLine($"EMAIL:{correo}");
                            sw.WriteLine("END:VCARD");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a vCard: " + ex.Message);
            }

            return filePath;  // Devolver la ruta del archivo generado
        }



        private void dgvContactos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            // Exportar a CSV y abrir el archivo
            string csvFilePath = ExportarCSV();  // Obtener la ruta del archivo CSV
            if (File.Exists(csvFilePath))
            {
                System.Diagnostics.Process.Start(csvFilePath);  // Abrir el archivo CSV con la aplicación predeterminada
            }

            // Exportar a vCard y abrir el archivo
            string vCardFilePath = ExportarVCard();  // Obtener la ruta del archivo vCard
            if (File.Exists(vCardFilePath))
            {
                System.Diagnostics.Process.Start(vCardFilePath);  // Abrir el archivo vCard con la aplicación predeterminada
            }

            // Mostrar un mensaje indicando que los archivos han sido exportados y abiertos
            MessageBox.Show($"Contactos exportados y abiertos:\nCSV: {csvFilePath}\nvCard: {vCardFilePath}");
        }
    }
}
