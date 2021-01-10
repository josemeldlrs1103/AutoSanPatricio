using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JuntaDirectiva
{
    class Operaciones
    {
        //Validación para existencia de directorio raíz
        public static void CrearDirectorio(string Ruta)
        {
            string Raíz = Path.Combine(Ruta, "Junta Directiva");
            if (!Directory.Exists(Raíz))
            {
                Directory.CreateDirectory(Raíz);
                CrearSubCarpetas(Raíz);
                CrearArchivos(Raíz);
            }
            else
            {
                CrearSubCarpetas(Raíz);
                CrearArchivos(Raíz);
            }
        }
        //Validación para existencia directorios para orden según la fecha de operación
        static void CrearSubCarpetas(string SubRuta)
        {
            DateTime FechaActual = DateTime.Now.Date;
            int Año = FechaActual.Year;
            string Mes = FechaActual.ToString("MMMM", CultureInfo.CreateSpecificCulture("es"));
            string Dirección = Path.Combine(SubRuta, Convert.ToString(Año));
            //Creación de la subcarpeta año
            if (!Directory.Exists(Dirección))
            {
                Directory.CreateDirectory(Dirección);
            }
            Dirección = Path.Combine(Dirección, Convert.ToString(Mes));
            //Creación de subcarpeta mes
            if (!Directory.Exists(Dirección))
            {
                Directory.CreateDirectory(Dirección);
            }
            //Creación de carpeta oculta con recursos a usar
            if (!Directory.Exists(Path.Combine(SubRuta, "Recursos")))
            {
                DirectoryInfo Direc = new DirectoryInfo(Path.Combine(SubRuta, "Recursos"));
                Direc.Create();
                Direc.Attributes = FileAttributes.Hidden;
            }
        }
        //Creación de los archivos
        static void CrearArchivos(string DirecciónArchivos)
        {
            string Listado = Path.Combine(DirecciónArchivos, "ListadoVecinos.txt");
            if (!File.Exists(Listado))
            {
                File.Create(Listado);
            }
            string[] ArchivosRecursos = { Path.Combine(DirecciónArchivos, "Recursos", "No.Recibo.txt"), Path.Combine(DirecciónArchivos, "Recursos", "DatosUsuario.txt"), Path.Combine(DirecciónArchivos, "Recursos", "FondosIniciales.txt") };
            for (int i = 0; i < 3; i++)
            {
                if (!File.Exists(ArchivosRecursos[i]))
                {
                    File.Create(ArchivosRecursos[i]);
                    File.SetAttributes(ArchivosRecursos[i], FileAttributes.Hidden);
                }
            }
            //Creación de hoja de excel
            Microsoft.Office.Interop.Excel.Application HojaFondos = new Microsoft.Office.Interop.Excel.Application();
            if (HojaFondos == null)
            {
                MessageBox.Show("Excel no está instalado, por lo que el programa no le será útil");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbook Libro = HojaFondos.Workbooks.Add();
            Libro.SaveAs(Path.Combine(DirecciónArchivos, "FondosSanPatricio.xlsx"));
            Libro.Close();
            HojaFondos.Quit();
        }
    }
}
