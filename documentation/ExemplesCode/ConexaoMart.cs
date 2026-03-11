using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ErwinAPI_AddIn;

namespace ErwinAPI_AddIn
{
   public class ConexaoMart
    {
        public string ConectaMart(string biblioteca, string modeloName)
        {
            string Biblioteca = biblioteca;
            string Modelo = modeloName;
            string Servidor = "192.168.30.15";
            string Porta = "18170";
            string Aplicacao = "MartServer";
            string Usuario = "Administrator";
            string Senha = "";
            string MartArquivo = "Mart://Mart/" + Biblioteca + "/" + Modelo + "?TRC=NO;SRV=" + Servidor + ";PRT=" + Porta + ";ASR=" + Aplicacao + ";UID=" + Usuario + ";PSW=" + Senha + ";";
            //Mart://Mart/ErwinAPI/eMovies_Alterado?TRC=NO;SRV=192.168.30.15;PRT=18170;ASR=MartServer;UID=Administrator;PSW=;
            return MartArquivo;
        }
    }       
}
 