using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes
{
    // Classe que armazena parametros (do usuario) de otmizacao do fluxo de potencia e seus resp. metodos. 
    public class ParametrosGUI
    {
        //variaveis 
        public string _hora;
        public string _ano;
        public string _mes;
        public string _tipoDia;
        public int _mesNum;
        public bool _fluxoDiarioNativo;
        public string _modoFluxo;
        public bool _allowForms;
        public string _pathRecursosPerm;
        public string _pathRaizGUI;
        public bool _usarTensoesBarramento;
        public string _tensaoSaidaBarUsuario;
        public bool _otmPorDemMax = false; // 
        public bool _otmPorEnergia = false; // booleanda que seta o tipo de otimizacao (Potencia ou Energia)
        private double _precisao = 500; // Precisao % do ajuste
        private float _incremento = 1.05F; // Passo % do incremento
        public bool _aproximaFluxoMensalPorDU = false; //booleana que seta decide o tipo de fluxo Mensal
        public bool _modoAnual = false;
        public double _loadMultAlternativo; // loadMult alternativo inserido pelo usuário

        // variaveis temporaria nao presentes na GUI
        public double loadMultDefault = 1; // LoadMult retornado qnd alim nao eh encontrado no arquivo de ajuste.
        public double loadMultAtual = 1;

        // Utilizado por feriadosWindow
        public string getNomeArquivoFeriadosCompleto()
        {
            return _pathRecursosPerm + "Feriados" + _ano + ".txt";
        }

        public void preencheVariaveis(MainWindow jan)
        {
            //Armazena os valores da interface, a fim de se manterem acessíveis por outro processo
            _hora = jan.horaTextBox.Text; //ok
            _ano = jan.anoTextBox.Text;
            _mes = jan.mesComboBox.Text;
            _tipoDia = jan.tipoDiaComboBox.Text;
            _mesNum = jan.mesComboBox.SelectedIndex + 1;
            _fluxoDiarioNativo = jan.fluxoDiarioNativoCheckBox.IsChecked.Value;
            _modoFluxo = jan.tipoFluxoComboBox.Text;
            _pathRecursosPerm = jan.caminhoPermTextBox.Text;
            _pathRaizGUI = jan.caminhoDSSTextBox.Text;
            _usarTensoesBarramento = jan.usarTensoesBarramentoCheckBox.IsChecked.Value;
            _tensaoSaidaBarUsuario = jan.tensaoSaidaBarTextBox.Text;
            _otmPorDemMax = jan.otimizaCheckBox.IsChecked.Value;
            _otmPorEnergia = jan.otimizaEnergiaCheckBox.IsChecked.Value;
            _aproximaFluxoMensalPorDU = jan.simplificaMesComDUCheckBox.IsChecked.Value;
            _allowForms = jan.AllowFormsCheckBox.IsChecked.Value;

            // preenche incremento
            setIncremento(jan.incrementoAjusteTextBox.Text);

            // transforma texto para double
            _loadMultAlternativo = Double.Parse(jan.loadMultAltTextBox.Text);
        }

        public void setIncremento(float i)
        {
            _incremento = i / 100 + 1;
        }

        public void setIncremento(string s)
        {
            float i = float.Parse(s, CultureInfo.InvariantCulture);
            _incremento = i / 100 + 1;
        }

        public float getIncremento()
        {
            return _incremento;
        }

        internal double getPrecisao()
        {
            return _precisao;
        }
    }
}
