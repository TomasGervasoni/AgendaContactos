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
using OfficeOpenXml;

namespace AgendaContactos
{
    public partial class frmBuscarContactos : Form
    {
        public frmBuscarContactos()
        {
            InitializeComponent();
        }
        // Método para recibir los datos de búsqueda y filtrar en el DataGridView
        public void FiltrarDatos(string nombre, string telefono, string correo)
        {
            // Ruta a la base de datos Access
            string databasePath = "BaseDatos\\Contactos.accdb";

            // Crear instancia de la conexión
            ConexionBD conexionBD = new ConexionBD(databasePath);

            try
            {
                // Abrir la conexión
                conexionBD.Abrir();

                // Construir la consulta SQL con las condiciones de búsqueda
                string query = "SELECT * FROM Contactos WHERE 1=1";
                if (!string.IsNullOrWhiteSpace(nombre))
                {
                    query += " AND Nombre LIKE '%' + ? + '%'";
                }
                if (!string.IsNullOrWhiteSpace(telefono))
                {
                    query += " AND Telefono LIKE '%' + ? + '%'";
                }
                if (!string.IsNullOrWhiteSpace(correo))
                {
                    query += " AND Correo LIKE '%' + ? + '%'";
                }

                // Crear el adaptador de datos
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, conexionBD.ObtenerConexion());

                // Añadir los parámetros de búsqueda
                if (!string.IsNullOrWhiteSpace(nombre))
                {
                    dataAdapter.SelectCommand.Parameters.AddWithValue("?", nombre);
                }
                if (!string.IsNullOrWhiteSpace(telefono))
                {
                    dataAdapter.SelectCommand.Parameters.AddWithValue("?", telefono);
                }
                if (!string.IsNullOrWhiteSpace(correo))
                {
                    dataAdapter.SelectCommand.Parameters.AddWithValue("?", correo);
                }

                // Llenar el DataTable con los datos filtrados
                DataTable dataTable = new DataTable();
                dataAdapter.Fill(dataTable);

                // Asignar los datos al DataGridView
                dgvReporteContactos.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al Buscar Contacto: " + ex.Message);
            }
            finally
            {
                // Cerrar la conexión
                conexionBD.Cerrar();
            }
        }

        private void EliminarContacto(string nombre, string apellido)
        {
            // Ruta a la base de datos Access
            string databasePath = "BaseDatos\\Contactos.accdb";
            ConexionBD conexionBD = new ConexionBD(databasePath);

            // Consulta SQL para eliminar el contacto
            string query = "DELETE FROM Contactos WHERE Nombre = ? AND Apellido = ?";

            try
            {
                // Abrir la conexión
                conexionBD.Abrir();

                // Crear el comando SQL
                using (OleDbCommand command = new OleDbCommand(query, conexionBD.ObtenerConexion()))
                {
                    // Añadir los parámetros al comando
                    command.Parameters.AddWithValue("?", nombre);
                    command.Parameters.AddWithValue("?", apellido);

                    // Ejecutar el comando
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al eliminar el contacto: " + ex.Message);
            }
            finally
            {
                // Cerrar la conexión
                conexionBD.Cerrar();
            }
        }


        private void btnEliminar_Click(object sender, EventArgs e)
        {
            // Verificar que haya una fila seleccionada
            if (dgvReporteContactos.SelectedRows.Count > 0)
            {
                // Confirmación de eliminación
                DialogResult result = MessageBox.Show("¿Está seguro de que desea eliminar este contacto?",
                                                      "Confirmar eliminación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Obtener el valor de la clave (ej: Nombre y Apellido) de la fila seleccionada
                    DataGridViewRow filaSeleccionada = dgvReporteContactos.SelectedRows[0];
                    string nombre = filaSeleccionada.Cells["Nombre"].Value.ToString();
                    string apellido = filaSeleccionada.Cells["Apellido"].Value.ToString();

                    // Eliminar el contacto de la base de datos
                    EliminarContacto(nombre, apellido);

                    // (Opcional) Eliminar la fila del DataGridView
                    dgvReporteContactos.Rows.Remove(filaSeleccionada);

                    // Mensaje de éxito
                    MessageBox.Show("Contacto eliminado correctamente.");
                }
            }
            else
            {
                MessageBox.Show("Por favor, seleccione un contacto para eliminar.");
            }
        }


        private void btnVolver_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private string GenerarReporteExcel()
        {
            // Establecer el contexto de la licencia para EPPlus
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Obtener la ruta de la carpeta "exel" dentro del directorio del proyecto
            string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string folderPath = Path.Combine(projectDirectory, "exel");

            // Crear la carpeta "exel" si no existe
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            // Definir la ruta para guardar el archivo Excel
            string filePath = Path.Combine(folderPath, "ReporteContactos.xlsx");

            // Crear paquete Excel
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Crear hoja de trabajo
                var worksheet = excel.Workbook.Worksheets.Add("Contactos");

                // Escribir cabeceras
                worksheet.Cells[1, 1].Value = "Nombre";
                worksheet.Cells[1, 2].Value = "Apellido";
                worksheet.Cells[1, 3].Value = "Teléfono";
                worksheet.Cells[1, 4].Value = "Correo";
                worksheet.Cells[1, 5].Value = "Categoría";

                // Conectar a la base de datos y obtener los contactos
                string databasePath = "BaseDatos\\Contactos.accdb";
                ConexionBD conexionBD = new ConexionBD(databasePath);

                try
                {
                    conexionBD.Abrir();
                    string query = "SELECT * FROM Contactos";
                    using (OleDbCommand command = new OleDbCommand(query, conexionBD.ObtenerConexion()))
                    {
                        using (OleDbDataReader reader = command.ExecuteReader())
                        {
                            int row = 2; // Empezar a escribir desde la fila 2

                            // Escribir datos en el Excel
                            while (reader.Read())
                            {
                                worksheet.Cells[row, 1].Value = reader["Nombre"].ToString();
                                worksheet.Cells[row, 2].Value = reader["Apellido"].ToString();
                                worksheet.Cells[row, 3].Value = reader["Telefono"].ToString();
                                worksheet.Cells[row, 4].Value = reader["Correo"].ToString();
                                worksheet.Cells[row, 5].Value = reader["Categoria"].ToString();
                                row++;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al generar el reporte: " + ex.Message);
                }
                finally
                {
                    conexionBD.Cerrar();
                }

                // Guardar el archivo Excel
                FileInfo excelFile = new FileInfo(filePath);
                excel.SaveAs(excelFile);
            }

            return filePath; // Devolver la ruta del archivo generado
        }






        private void btnExportar_Click(object sender, EventArgs e)
        {

            // Generar el reporte y obtener la ruta del archivo
            string reportePath = GenerarReporteExcel();

            // Mostrar el archivo en Excel
            if (File.Exists(reportePath))
            {
                System.Diagnostics.Process.Start(reportePath); // Abre el archivo con la aplicación predeterminada (normalmente Excel)
            }

            // Mostrar mensaje con la ruta del archivo generado
            MessageBox.Show($"Reporte generado: {reportePath}");
        }




        private void btnModificar_Click(object sender, EventArgs e)
        {
            // Verificar que haya una fila seleccionada
            if (dgvReporteContactos.SelectedRows.Count > 0)
            {
                // Obtener los datos de la fila seleccionada
                DataGridViewRow filaSeleccionada = dgvReporteContactos.SelectedRows[0];
                string nombre = filaSeleccionada.Cells["Nombre"].Value.ToString();
                string apellido = filaSeleccionada.Cells["Apellido"].Value.ToString();
                string telefono = filaSeleccionada.Cells["Telefono"].Value.ToString();
                string correo = filaSeleccionada.Cells["Correo"].Value.ToString();
                string categoria = filaSeleccionada.Cells["Categoria"].Value.ToString();

                // Crear instancia del formulario de actualización
                frmActualizarContacto formActualizar = new frmActualizarContacto();

                // Pasar los datos al formulario de actualización
                formActualizar.CargarDatosContacto(nombre, apellido, telefono, correo, categoria);

                // Mostrar el formulario de actualización
                formActualizar.ShowDialog();

                // (Opcional) Refrescar los datos del DataGridView después de la actualización
                FiltrarDatos(txtNombre.Text, txtTelefono.Text, txtCorreo.Text);
            }
            else
            {
                MessageBox.Show("Por favor, seleccione un contacto para modificar.");
            }
        }

        private void frmBuscarContactos_Load(object sender, EventArgs e)
        {

        }

        private void btnBuscarContactos_Click(object sender, EventArgs e)
        {
            // Obtener los valores de los TextBox para la búsqueda
            string nombre = txtNombre.Text;
            string telefono = txtTelefono.Text;
            string correo = txtCorreo.Text;

            // Filtrar los contactos
            FiltrarDatos(nombre, telefono, correo);
        }

        private void dgvReporteContactos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
    
}
