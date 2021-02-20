using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using frmWindows = AmgSistemas.Framework.WindowsForms;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            List<Empresa> Items = new List<Empresa>();

        Items.Add(new Empresa() {Codigo = 1, Nome = "Anselmo", CodAcesso = "1", Filiais = new List<Filial>()});
        Items.Add(new Empresa() {Codigo = 1, Nome = "Anselmo", CodAcesso = "1", Filiais = new List<Filial>()});

        List<frmWindows.Item> teste = frmWindows.Util.ConverterItems(Items, "Codigo", "Nome");

        }
    }
    public class Empresa
{
    public string Identificador;
    public string Nome;
    public int Codigo;
    public string CodAcesso;
    public List<Filial> Filiais;
}


public class Filial
{
    public string Identificador;
    }

}
