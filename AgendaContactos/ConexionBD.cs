using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AgendaContactos
{
    internal class ConexionBD
    {
        private string connectionString;
        private OleDbConnection connection;

        public ConexionBD(string databasePath)
        {
            // Define la cadena de conexión. Asegúrate de reemplazar "NombreDeTuBD.accdb" con el nombre real de tu base de datos.
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databasePath + ";Persist Security Info=False;";
            connection = new OleDbConnection(connectionString);
        }

        // Abre la conexión a la base de datos
        public void Abrir()
        {
            try
            {
                if (connection.State == ConnectionState.Closed)
                {
                    connection.Open();
                    //MessageBox.Show("apertura");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al abrir la conexión: " + ex.Message);
            }
        }

        // Cierra la conexión a la base de datos
        public void Cerrar()
        {
            try
            {
                if (connection.State == ConnectionState.Open)
                {
                    connection.Close();
                    //MessageBox.Show("cierre");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al cerrar la conexión: " + ex.Message);
            }
        }

        // Método para obtener la conexión
        public OleDbConnection ObtenerConexion()
        {
            return connection;
        }
    }
}
