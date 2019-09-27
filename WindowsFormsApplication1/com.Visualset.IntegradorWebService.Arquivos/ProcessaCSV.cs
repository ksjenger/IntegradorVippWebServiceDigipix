using IntegradorWebService.WSVIPP;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IntegradorWebService.ExcelServices
{
    class ProcessaTxt
    {
        public static List<Postagem> ListaDePostagem(string path, Form1 frm)
        {
            #region Recupera a formatação da planilha do Settings.settings
            List<FormatacaoPlanilha> lFormatacao = new List<FormatacaoPlanilha>();
            lFormatacao = FormatacaoPlanilha.ListarFormatacao();
            #endregion

            List<Postagem> lVipp = new List<Postagem>();
            Destinatario oDestinatario = new Destinatario();
            VolumeObjeto[] oVolumeObjetos = new VolumeObjeto[] {};
            Servico oServico = new Servico();

            //atribui a uma matriz os valores inseridos no arquivo path
            var list = File.ReadAllLines(path).Select(a => a.Split(';')).ToList();

            #region Lista de Formatação
            //For para percorrer a lista de Formatação

            foreach (FormatacaoPlanilha lista in lFormatacao)
            {
                //for que percorre cada linha da matriz
                int cont = 0;
                foreach(var linha in list)
                {
                    cont++;
                    string atributo = lista.NomeAtributo;
                    int coluna = lista.Coluna;

                    if(coluna == cont++)
                    {
                                                 
                    }
                        

                }


            }//fim do For da Lista de Formatacao

            #endregion


            return lVipp;
        }
    }

}
