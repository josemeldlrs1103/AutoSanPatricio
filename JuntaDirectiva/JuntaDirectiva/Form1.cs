﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JuntaDirectiva
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //Creación de carpeta y archivos iniciales en el escritorio 
            Operaciones.CrearDirectorio(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
        }
    }
}
