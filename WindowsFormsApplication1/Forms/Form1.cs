using System;
using System.Collections.Generic;
using System.Windows.Forms;
using IntegradorWebService.Services;
using IntegradorWebService.WSVIPP;
using IntegradorWebService.ExcelServices;
using IntegradorWebService.Rest;

namespace IntegradorWebService
{
    public partial class Form1 : Form
    {

        List<Postagem> lVipp = new List<Postagem>();
        Rootobject lPerfil = new Rootobject();

        public static string path;
        public static string nomeArquivo;
        public static string caminhoArquivo;

        public Form1(string usuario, string senha)
        {
            InitializeComponent();
            this.Text = "Importador Visual Personalizado - Versão: " + Application.ProductVersion;
            Cursor = default;
            btnEnviar.Enabled = false;
            lPerfil = RestPerfilImportacao.ProcessaListaPerfil(usuario, senha);
            comboPerfil.Items.Add("Selecione o Perfil");
            comboPerfil.SelectedIndex = 0;
            for (int i = 0; i < lPerfil.Data.Length; i++)
            {
                comboPerfil.Items.Add(lPerfil.Data[i].IdPerfil + " - " + lPerfil.Data[i].NomePerfil);
            }

        }


        private void Button2_Click(object sender, EventArgs e)
        {

            int id = comboPerfil.SelectedIndex;

            if (id == 0)
            {
                MessageBox.Show("Selecione o perfil de importação", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {

                Login.Operfil.IdPerfil = lPerfil.Data[id-1].IdPerfil;

                #region Chama o metodo para Postar Objeto
                VIPP.PostarObjeto.Postagem(lVipp, this);
                #endregion

                labelProgresso.Text = "Salvando o arquivo processado...";
                GravaRetornoExcel.GravaRetorno();
                MessageBox.Show("Importação finalizada", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                path = null;
                labelPath.Text = "";
                labelProgresso.Text = "";
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {



        }

        private void ComboPerfil_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(path == null)
            {
                btnSelecione.Focus();
            }
            else
            {
                btnEnviar.Focus();
            }
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            btnEnviar.Enabled = false;
            #region Abre o Arquivo
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                path = openFileDialog.FileName;
                nomeArquivo = System.IO.Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                caminhoArquivo = System.IO.Path.GetDirectoryName(openFileDialog.FileName);
                labelPath.Text = path;
                labelProgresso.Text = "Importando o Arquivo";
                lVipp = ProcessaPlanilha.ListaDePostagem(path, this);
                labelProgresso.Text = "Arquivo importado!";
                btnEnviar.Enabled = true;
                comboPerfil.Focus();
            }
        }
        #endregion

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.visualset.com.br");
            System.Diagnostics.Process.Start("http://vipp.visualset.com.br");
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
