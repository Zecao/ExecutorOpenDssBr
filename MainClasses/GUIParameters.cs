using ExecutorOpenDSS.AuxClasses;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Globalization;

namespace ExecutorOpenDSS.MainClasses
{
    // Class to store user parameters
    public class GUIParameters
    {
        //variaveis 
        public string _hora;
        public string _ano = "2022";
        public string _mes;
        public string _tipoDia;
        private int _mesNum;
        public string _tipoFluxo;
        public bool _allowForms;
        public string _pathRecursosPerm;
        public string _pathRaizGUI;
        public bool _usarTensoesBarramento;
        public string _tensaoSaidaBarUsuario;
        public bool _otmPorDemMax = false; // 
        public bool _otmPorEnergia = false; // boolean to set the type of optimization (Power or Energy)
        private readonly double _precisao = 500; // adjustment accuracy %
        private float _incremento = 1.05F; // increment step %
        private bool _aproximaFluxoMensalPorDU = false; // booleana to set the type of power flow 
        public bool _modoAnual = false;
        public bool _modoHorario = false;
        public double _loadMultAlternativo; // loadMult alternativo inserido pelo usuário
        public bool _usarLoadMult = true;

        public string _mesAbrv3letras;
        private readonly string _lstAlim = "lstAlimentadores.m"; // TODO 

        // expander parameters
        public ExpanderParameters _expanderPar;

        //
        public string GetCodTipoDia()
        {
            //default;
            string tipoDia = "DU";
            // so faz associacaose tipo fluxo for igual da daily ou hourly 
            if (_tipoFluxo.Equals("Daily") || _tipoFluxo.Equals("Hourly"))
            {
                switch (_tipoDia)
                {
                    case "Sábado":
                        tipoDia = "SA";
                        break;
                    case "Domingo":
                        tipoDia = "DO";
                        break;
                }
            }
            return tipoDia;
        }

        // set Mes 
        public void SetMes(int mes)
        {
            _mesNum = mes;

            //atualiza abreviatua mes
            _mesAbrv3letras = TipoDiasMes.GetMesAbrv(mes);
        }

        public int GetMes()
        {
            return _mesNum;
        }

        // get arquivo lista Alimentadores 
        public string GetArqLstAlimentadores()
        {
            return _pathRecursosPerm + _lstAlim;
        }

        // Utilizado por feriadosWindow
        public string GetNomeArquivoFeriadosCompleto()
        {
            return _pathRecursosPerm + "Feriados" + _ano + ".txt";
        }

        // 
        public void CopiaVariaveis(MainWindow jan)
        {
            //Armazena os valores da interface
            _hora = jan.horaTextBox.Text;
            _ano = jan.anoTextBox.Text;
            _mes = jan.mesComboBox.Text;

            _tipoDia = jan.tipoDiaComboBox.Text;
            _mesNum = jan.mesComboBox.SelectedIndex + 1;
            _mesAbrv3letras = _mes.Substring(0, 3); //OBS: Alternativa _mesAbrv3letras = TipoDiasMes.getMesAbrv(_mesNum);
            _tipoFluxo = jan.tipoFluxoComboBox.Text;
            _pathRecursosPerm = jan.caminhoPermTextBox.Text;
            _pathRaizGUI = jan.caminhoDSSTextBox.Text;
            _usarTensoesBarramento = jan.usarTensoesBarramentoCheckBox.IsChecked.Value;
            _tensaoSaidaBarUsuario = jan.tensaoSaidaBarTextBox.Text;
            _otmPorDemMax = jan.otimizaCheckBox.IsChecked.Value;
            _otmPorEnergia = jan.otimizaEnergiaCheckBox.IsChecked.Value;
            _aproximaFluxoMensalPorDU = jan.simplificaMesComDUCheckBox.IsChecked.Value;

            // preenche incremento
            SetIncremento(jan.incrementoAjusteTextBox.Text);

            // transforma texto para double
            _loadMultAlternativo = Double.Parse(jan.loadMultAltTextBox.Text);

            //
            _expanderPar = new ExpanderParameters(jan);
        }

        public void SetIncremento(string s)
        {
            float i = float.Parse(s, CultureInfo.InvariantCulture);
            _incremento = i / 100 + 1;
        }

        public float GetIncremento()
        {
            return _incremento;
        }

        public double GetPrecisao()
        {
            return _precisao;
        }

        public bool GetAproximaFluxoMensalPorDU()
        {
            return _aproximaFluxoMensalPorDU;
        }
    }
}
