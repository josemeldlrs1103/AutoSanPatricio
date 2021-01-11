using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

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
                HojaMes(Raíz);
            }
            else
            {
                CrearSubCarpetas(Raíz);
                CrearArchivos(Raíz);
                HojaMes(Raíz);
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
            //Instancia de objeto aplicación para excel
            _Excel.Application HojaFondos = new _Excel.Application();
            if (HojaFondos == null)
            {
                MessageBox.Show("Excel no está instalado, por lo que el programa no le será útil");
                return;
            }
            _Excel.Workbook Libro = HojaFondos.Workbooks.Add();
            if (!File.Exists(Path.Combine(DirecciónArchivos, "FondosSanPatricio.xlsx")))
            {
                Libro.SaveAs(Path.Combine(DirecciónArchivos, "FondosSanPatricio.xlsx"));
            }
            Libro.Close();
            HojaFondos.Quit();
        }
        //Creación de la hoja del mes en el libro de excel
        static void HojaMes(string RutaArchivo)
        {
            //Instancia para el manejo de libro del documento creado
            _Excel.Application DocumentoAbierto = new _Excel.Application();
            _Excel.Workbook LibroAbierto = DocumentoAbierto.Workbooks.Open(Path.Combine(RutaArchivo, "FondosSanPatricio.xlsx"));
            //Crear una cadena con el nombre de la hoja para el mes actual
            DateTime Fecha = DateTime.Now.Date;
            string NombreHoja = Fecha.ToString("MMMM", CultureInfo.CreateSpecificCulture("es")) + " " + Convert.ToString(Fecha.Year);
            //Bandera para la validación de la Hoja
            bool HojaExiste = false;
            foreach(Worksheet Hoja in LibroAbierto.Worksheets)
            {
                if(Hoja.Name == NombreHoja)
                {
                    HojaExiste = true;
                    break;
                }
            }
            //Crear una nueva hoja y cambiar el nombre de la misma
            if(!HojaExiste)
            {
                LibroAbierto.Worksheets.Add();
                LibroAbierto.Worksheets[LibroAbierto.Worksheets.Count].Name = NombreHoja;
            }
            int IndiceHoja = 0;
            foreach(Worksheet HojaSelecta in LibroAbierto.Worksheets)
            {
                if(HojaSelecta.Name.Contains("Hoja"))
                {
                    IndiceHoja = HojaSelecta.Index;
                    LibroAbierto.Sheets[IndiceHoja].Delete();
                }
            }
            LibroAbierto.Save();
            LibroAbierto.Close();
            DocumentoAbierto.Quit();
        }
    }
}
