using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Testes
{
    public partial class Form1 : Form
    {
        private List<Testando> objTeste;
        public Form1()
        {
            InitializeComponent();

            objTeste = new List<Testando>();
            objTeste.Add(new Testando()
            {
                Codigo = 1,
                Descricao = "Anselmo",
                Identificador = Guid.NewGuid().ToString()
            });

            List<AmgSistemas.Framework.WindowsForms.Item> objItems = AmgSistemas.Framework.WindowsForms.Util.ConverterItems(objTeste, "Identificador", "Descricao");
          listBox1 =   AmgSistemas.Framework.WindowsForms.Util.PreencherListBox(listBox1, objItems);
        }
    }

    public class Testando
    {
        public string Identificador { get; set; }
        public Int32 Codigo { get; set; }
        public string Descricao { get; set; }
    }
}
