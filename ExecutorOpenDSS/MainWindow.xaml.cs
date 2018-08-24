using ExecutorOpenDSS.Classes;
using ExecutorOpenDSS.Classes_Auxiliares;
using ExecutorOpenDSS.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExecutorOpenDSS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///

   
    public partial class MainWindow : Window
    {
        //Variáveis Globais da classe
        public System.Windows.Threading.Dispatcher _tbdDispatcher;
        public System.Windows.Threading.Dispatcher _ebDispatcher;

        // variaveis da GUI
        public ParametrosGUI _parGUI = new ParametrosGUI();
        public ParamAvancados _parAvan = new ParamAvancados();

        // variaveis classe 
        public bool _cancelarExecucao;
        public string _textoDisplay;
        public bool _fimExecucao = false;
        public string _logName;

        public DateTime _inicio;

        public int _indiceArq;

        /*
        public void Close()
        {
            //Finalização do processo
            FinalizaProcesso();

            this.Close();
        }*/

        //Função principal
        public MainWindow()
        {
            InitializeComponent();

            //Inicializa os valores da interface
            InicializaValores();
        }

        //Função que exibe as mensagens na TextBox e as grava em um arquivo de log
        public void Disp(string mensagem, string log = "")
        {            
            //Pega o texto do display. Como a interface roda em outro processo, é necessária a utilização
            //de função delegada, dispatcher e invoke
            MainWindow.GetDisplayDelegate getDisplay = new MainWindow.GetDisplayDelegate(this.GetDisplay);

            //
            string str = this._tbdDispatcher.Invoke(getDisplay).ToString(); 

            //Verifica se o display está em branco.
            if (str.Equals(""))
            {
                //Adiciona a mensagem
                str = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + mensagem;
            }
            else
            {
                //Caso o display não esteja em branco, pula uma linha e adiciona a mensagem
                str = str + "\n" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ": " + mensagem;
            }

            //Escreve no display da tela.Como a interface roda em outro processo, é necessária a utilização
            //de função delegada, dispatcher e invoke
            MainWindow.UpdateDisplayDelegate update = new MainWindow.UpdateDisplayDelegate(this.UpdateDisplay);
            this._tbdDispatcher.BeginInvoke(update, str);
        }

        //Função chamada quando o botão "Executar" é clicado
        private void ExecutaButton_Click(object sender, RoutedEventArgs e)
        {
            //Desabilita a interface
            StatusUI(false);

            //Inicializa o log
            InicializaLog();
            
            //TODO 
            _indiceArq = 0;
           
            //limpa display
            display.Text = "";
            
            // preenche variaveis da GUI antes da exeucao
            _parGUI.preencheVariaveis(this);

            // Inicializa um novo processo em background.
            // O código é executado em um novo processo para manter a interface responsiva durante a execução
            _tbdDispatcher = display.Dispatcher;
            _ebDispatcher = this.Dispatcher;

            // cria obj worker para execucao do fluxo em background
            BackgroundWorker worker = new BackgroundWorker();

            //
            worker.DoWork += worker_ExecutaFluxo;

            // Roda worker_ExecutaFluxo em background
            worker.RunWorkerAsync();

            //
            GravaConfig();
        }

        // get arquivo lista Alimentadores 
        private string getArqLstAlimentadores()
        { 
            return _parGUI._pathRaizGUI + "Recursos\\lstAlimentadores.m"; 
        }

        // Executa Fluxo potencia
        void worker_ExecutaFluxo(object sender, DoWorkEventArgs e)
        {
            //
            _inicio = DateTime.Now;

            //Mensagem de Início
            Disp("Início");

            //Lê os alimentadores e armazena a lista de alimentadores 
            List<string> alimentadores = AlimentadoresCemig.getTodos( getArqLstAlimentadores() );
          
            // instancia classe de parametros Gerais
            ParamGeraisDSS paramGerais = new ParamGeraisDSS(this);

            // instancia classe ExecutaFluxo
            ExecutaFluxo executaFluxoObj = new ExecutaFluxo(this, paramGerais);

            switch (_parGUI._modoFluxo)
            {             
                //Executa o Fluxo Snap
                case "Snap":

                    executaFluxoObj.executaSnap(alimentadores);

                    break;

                //Executa o fluxo diário
                case "Daily":
                     
                    executaFluxoObj.executaDiario(alimentadores);

                    break;

                //Executa o fluxo diário
                case "Hourly":

                    executaFluxoObj.executaDiario(alimentadores);
                    break;

                //Executa o fluxo mensal
                case "Monthly":

                    executaFluxoObj.executaMensal(alimentadores);
                    
                    break;

                //Executa o fluxo mensal
                case "Yearly":

                    executaFluxoObj.executaAnual(alimentadores);

                    break;

                default:
                    break;
            }

            //Finalização do processo
            FinalizaProcesso();
        }

        //Botão para cancelar a execução 
        private void CancelaButton_Click(object sender, RoutedEventArgs e)
        {
            _cancelarExecucao = true;
            CancelaButton.IsEnabled = false;
        }

        // Se o Fluxo for Snap ou Monthly, desabilita a combobox do tipo de dia
        private void tipoFluxoCB_LostFocus(object sender, RoutedEventArgs e)
        {
            switch (tipoFluxoComboBox.Text)
            {
                case "Snap":

                    //1. desabilita a combobox do tipo de dia
                    tipoDiaComboBox.IsEnabled = false;
                  
                    //2. desabilita check-box RELACIONADOS otimiza LoadMult Energia
                    otimizaEnergiaCheckBox.IsEnabled = false;
                    otimizaEnergiaCheckBox.IsChecked = false;
                    simplificaMesComDUCheckBox.IsEnabled = false;
                    simplificaMesComDUCheckBox.IsChecked = false;

                    //3. desabilita checkBox de escolha entre o fluxo diario nativo ou 24vezes snap
                    fluxoDiarioNativoCheckBox.IsEnabled = false;

                    //4. desabilita checkBox de relatorios
                    horaTextBox.IsEnabled = false;

                    //5. habilita combo box do tipo dia
                    tipoDiaComboBox.IsEnabled = false;

                    //6. desabilita otimiza energia
                    otimizaEnergiaCheckBox.IsEnabled = false;

                    //8. desabilita otimiza Snap 
                    otimizaCheckBox.IsEnabled = true;
                    
                    break;

                case "Hourly":

                    //2. desabilita check-box RELACIONADOS otimiza LoadMult Energia
                    otimizaEnergiaCheckBox.IsEnabled = false;
                    otimizaEnergiaCheckBox.IsChecked = false;
                    simplificaMesComDUCheckBox.IsEnabled = false;
                    simplificaMesComDUCheckBox.IsChecked = false;

                    //3. desabilita checkBox de escolha entre o fluxo diario nativo ou 24vezes snap
                    fluxoDiarioNativoCheckBox.IsEnabled = false;

                    //4. habilita o textBox da hora
                    horaTextBox.IsEnabled = true;

                    //5. habilita combo box do tipo dia
                    tipoDiaComboBox.IsEnabled = true;

                    //6. desabilita otimiza energia
                    otimizaEnergiaCheckBox.IsEnabled = false;

                    //8. desabilita otimiza Snap 
                    otimizaCheckBox.IsEnabled = false;

                    break;

                case "Daily":

                    //2. desabilita check-box RELACIONADOS otimiza LoadMult Energia
                    otimizaEnergiaCheckBox.IsEnabled = false;
                    otimizaEnergiaCheckBox.IsChecked = false;
                    simplificaMesComDUCheckBox.IsEnabled = false;
                    simplificaMesComDUCheckBox.IsChecked = false;

                    //4. desabilita checkBox de relatorios
                    horaTextBox.IsEnabled = false;

                    //6. habilita combo box do tipo dia
                    tipoDiaComboBox.IsEnabled = true;

                    //7. desabilita otimiza energia
                    otimizaEnergiaCheckBox.IsEnabled = false;

                    //8. desabilita otimiza Snap 
                    otimizaCheckBox.IsEnabled = false;

                    break;

                case "Monthly":

                    //1. desabilita a combobox do tipo de dia
                    tipoDiaComboBox.IsEnabled = false;

                    //2.1 desabilita o check-box otimizacao potencia
                    otimizaCheckBox.IsEnabled = false;
                    otimizaCheckBox.IsChecked = false;

                    //3. desabilita checkBox de escolha entre o fluxo diario nativo ou 24vezes snap
                    fluxoDiarioNativoCheckBox.IsEnabled = false;

                    //4. desabilita checkBox de relatorios
                    horaTextBox.IsEnabled = false;

                    //5. habilita combo box do mes
                    //mesComboBox.Text = "Janeiro";
                    mesComboBox.IsEnabled = true;

                    //6. desabilita combo box do tipo dia
                    //tipoDiaComboBox.Text = "Dia Útil";
                    tipoDiaComboBox.IsEnabled = false;

                    //8. desabilita otimiza Snap 
                    otimizaCheckBox.IsEnabled = false;

                    //7. habilita otimiza energia
                    otimizaEnergiaCheckBox.IsEnabled = true;

                    break;

                case "Yearly":

                    //1. desabilita a combobox do tipo de dia
                    tipoDiaComboBox.IsEnabled = false;

                    //2.1 desabilita o check-box otimizacao potencia
                    otimizaCheckBox.IsEnabled = false;
                    otimizaCheckBox.IsChecked = false;

                    //3. desabilita checkBox de escolha entre o fluxo diario nativo ou 24vezes snap
                    fluxoDiarioNativoCheckBox.IsEnabled = false;

                    //4. desabilita checkBox de relatorios
                    horaTextBox.IsEnabled = false;

                    //5. ajusta combo box do mes
                    mesComboBox.Text = "Janeiro"; // TODO
                    mesComboBox.IsEnabled = false;

                    //6. ajusta combo box do tipo dia
                    tipoDiaComboBox.IsEnabled = false;

                    //7. ajusta text box da hora
                    horaTextBox.IsEnabled = false;

                    //8. desabilita otimiza Snap 
                    otimizaCheckBox.IsEnabled = false;

                    //7. habilita otimiza energia
                    otimizaEnergiaCheckBox.IsEnabled = true;

                    break;

                //Caso contrário, habilita a combobox do tipo de dia
                default:

                    break;
            }
        }

        /*
        //TODO 
        private void HabilitaInterface()
        {
            //habilita checkBox de escolha entre o fluxo diario nativo ou 24vezes snap
            fluxoDiarioNativoCheckBox.IsEnabled = true;
        }*/
        
        //Verifica se o ano é válido
        private void anoTB_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                //Tenta um parse para inteiro na caixa de texto de ano
                int ano = Int32.Parse(anoTextBox.Text);

                //Se o ano for menor ou igual a 0 ou maior que 3000, lança uma exceção de índice fora do intervalo
                if(ano<=0||ano>3000){
                    throw new IndexOutOfRangeException();
                }

                //Define a cor do texto para preto, caso não tenham ocorrido erros acima
                anoTextBox.Foreground = Brushes.Black;
                anoTextBox.BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0xE6, 0xE3, 0xE3));
            }
            //Caso ocorra erros no try acima:
            catch
            {
                //Define a cor do texto da caixa como vermelho
                anoTextBox.Foreground = Brushes.Red;
                anoTextBox.BorderBrush = Brushes.Red;

                //Exibe uma caixa de mensagem avisando que o ano é inválido
                MessageBox.Show("Ano inválido","Erro!");
            }
        }
              
        //Função executada quando a TextBox caminhoDSS perde o foco.
        //Valida o caminho.
        private void caminhoDSSTB_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!caminhoDSSTextBox.Text.Equals(""))
            {
                //Verifica se o caminhoDSS termina com "\". Caso contrário, acrescenta
                caminhoDSSTextBox.Text = caminhoDSSTextBox.Text.Last() == '\\' ? caminhoDSSTextBox.Text : caminhoDSSTextBox.Text + '\\';
            }
            
            //Verifica se o caminho existe
            if (Directory.Exists(caminhoDSSTextBox.Text))
            {
                //Caso exista, muda a cor do texto para preto
                caminhoDSSTextBox.Foreground = Brushes.Black;
                caminhoDSSTextBox.BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0xE6, 0xE3, 0xE3));
            }
            else
            {
                //Caso não exista, muda a cor do texto para vermelho
                caminhoDSSTextBox.Foreground = Brushes.Red;
                caminhoDSSTextBox.BorderBrush = Brushes.Red;

                //Exibe mensagem que o caminho não existe
                MessageBox.Show("Caminho dos arquivos DSS não existe", "Erro!");
            }
        }

        //Abre caixa de dialogo para seleção do caminho dos arquivos DSS
        private void caminhoDSSBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();

            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                caminhoDSSTextBox.Text = dialog.SelectedPath;
                caminhoDSSTB_LostFocus(caminhoDSSTextBox, new RoutedEventArgs());
            }
        }

        //Seleciona todo o conteúdo da caixa de texto, quando realiza-se um clique duplo
        private void TextBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            (sender as TextBox).SelectAll();
        }

        //Valida valores de caixas de texto cujo tipo é ponto flutuante
        private void floatTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox sendert = (sender as TextBox);

            sendert.Text = sendert.Text.Replace(',', '.');

            try
            {
                //Tenta um parse para float na caixa de texto
                float valor = float.Parse(sendert.Text, CultureInfo.InvariantCulture);

                //Define a cor do texto para preto, caso não tenham ocorrido erros acima
                sendert.Foreground = Brushes.Black;
                sendert.BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0xE6, 0xE3, 0xE3));
            }
            //Caso ocorra erros no try acima:
            catch
            {
                //Define a cor do texto da caixa como vermelho
                sendert.Foreground = Brushes.Red;
                sendert.BorderBrush = Brushes.Red;

                //Exibe uma caixa de mensagem avisando que o ano é inválido
                MessageBox.Show("Valor inválido inserido", "Erro!");
            }
        }
        
        //Desabilita/habilita a interface
        public void StatusUI(bool status)
        {
            ExecutaButton.IsEnabled = status;
            anoTextBox.IsEnabled = status;
            mesComboBox.IsEnabled = status;
            tipoFluxoComboBox.IsEnabled = status;
            caminhoDSSTextBox.IsEnabled = status;
            caminhoPermTextBox.IsEnabled = status;
            caminhoDSSBrowserButton.IsEnabled = status;
            otimizaCheckBox.IsEnabled = status;          
            usarTensoesBarramentoCheckBox.IsEnabled = status;
            tensaoSaidaBarTextBox.IsEnabled = status;
            otimizaEnergiaCheckBox.IsEnabled = status;
            fluxoDiarioNativoCheckBox.IsEnabled = status;
            simplificaMesComDUCheckBox.IsEnabled = status;
            incrementoAjusteTextBox.IsEnabled = status;
            loadMultAltTextBox.IsEnabled = status;

            // TODO ?
            if (status)
            {
                tipoFluxoCB_LostFocus(tipoFluxoComboBox,new RoutedEventArgs());
            }
            CancelaButton.IsEnabled = !status;
        }        
        
        //Finalização do processo
        public void FinalizaProcesso()
        {
            if (_cancelarExecucao)
            {   
                Disp("Execução abortada.");
            }
            else
            {
                Disp("Fim da execução.");
            }

            _fimExecucao = true;
            _cancelarExecucao = false;
            GravaLog();

            // data final para calculo do tempo de execucao
            DateTime fim = DateTime.Now;

            Disp("Tempo total de execução: " + (fim - _inicio).ToString());

            //Reabilita a interface. 
            //Como a interface está em outro processo, é necessário utilizar um Dispatcher
            SetButtonDelegate setar = new SetButtonDelegate(SetButton);
            _ebDispatcher.BeginInvoke(setar);
        }
        
        //Inicializa valores da interface
        private void InicializaValores()
        {
            _cancelarExecucao = false;

            string arqConfig = AppDomain.CurrentDomain.BaseDirectory + "Configuracoes.xml";
            Configuracoes.GetConfiguracoes(this, arqConfig);
            
            // TODO ?
            tipoFluxoCB_LostFocus(tipoFluxoComboBox, new RoutedEventArgs());
        }
        
        //Grava configurações
        private void GravaConfig()
        {
            string arqConfig = AppDomain.CurrentDomain.BaseDirectory + "Configuracoes.xml";
            Configuracoes.SetConfiguracoes(this, arqConfig);
        }

        //Grava o conteúdo da caixa de texto "display" em um arquivo de log nomeado com a data e hora atuais
        private void GravaLog()
        {
            GetDisplayDelegate getDisplay = new GetDisplayDelegate(GetDisplay);
            using (StreamWriter file = new StreamWriter(_logName, true))
            {
                string str = _tbdDispatcher.Invoke(getDisplay).ToString();
                file.Write(str);
            }
        }
        
        // Inicializa o arquivo de log "apendando" uma linha com um separador, dividindo o log anterior do atual
        private void InicializaLog()
        {
            //Define o nome do arquivo de log
            _logName = AppDomain.CurrentDomain.BaseDirectory + "Log.txt";
            
            //Apaga o arquivo de log existente
            ArqManip.SafeDelete(_logName);            
        }     
                
        /////////////////////////////////////////////////////////////////////////////////////////////////////
        //Funções para chamadas inter-processos//////////////////////////////////////////////////////////////
 
        //Funções para setar o texto do display
        public delegate void UpdateDisplayDelegate(string texto);

        public void UpdateDisplay(string texto)
        {
            display.Text = texto;
            display.ScrollToEnd();
        }

        //Funções para pegar o texto do display
        public delegate string GetDisplayDelegate();

        //
        public string GetDisplay()
        {
            return display.Text;
        }

        //Funções para habilitar/desabilitar a interface
        public delegate void SetButtonDelegate();

        //
        public void SetButton()
        {
            StatusUI(true);
        }

        //Exibir caixa de mensagem
        public delegate bool MensagemDelegate(string texto, string titulo="");

        //
        public bool Mensagem(string texto, string titulo = "Aviso")
        {
            MessageBoxButton botoes = MessageBoxButton.YesNo;
            MessageBoxResult resultado =MessageBox.Show(texto, titulo, botoes);
            if (resultado == MessageBoxResult.Yes)
            {
                return true;
            }
            else
            {
                return false;
            }            
        }

        //
        private void otimizaCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            // desabilita outro otimiza
            otimizaEnergiaCheckBox.IsChecked = false;

            // habilita Interfaces
            incrementoAjusteTextBox.IsEnabled = true;
            loadMultAltTextBox.IsEnabled = true;
        }

        //
        private void otimizaCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            // habilita Interfaces
            incrementoAjusteTextBox.IsEnabled = false;
            loadMultAltTextBox.IsEnabled = false;
        }

        private void otimizaEnergiaCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            // desabilita outro otimiza
            otimizaCheckBox.IsChecked = false;            

            // habilita Interfaces
            incrementoAjusteTextBox.IsEnabled = true;
            simplificaMesComDUCheckBox.IsEnabled = true;
            loadMultAltTextBox.IsEnabled = true;
        }

        private void otimizaEnergiaCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            // desabilita Interfaces
            incrementoAjusteTextBox.IsEnabled = false;
            simplificaMesComDUCheckBox.IsEnabled = false;
            loadMultAltTextBox.IsEnabled = false;
        }

        // Caminho
        private void caminhoTB_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!caminhoPermTextBox.Text.Equals(""))
            {
                //Verifica se o caminhoDSS termina com "\". Caso contrário, acrescenta
                caminhoPermTextBox.Text = caminhoPermTextBox.Text.Last() == '\\' ? caminhoPermTextBox.Text : caminhoPermTextBox.Text + '\\';
            }

            //Verifica se o caminho existe
            if (Directory.Exists(caminhoPermTextBox.Text))
            {
                //Caso exista, muda a cor do texto para preto
                caminhoPermTextBox.Foreground = Brushes.Black;
                caminhoPermTextBox.BorderBrush = new SolidColorBrush(Color.FromArgb(0xFF, 0xE6, 0xE3, 0xE3));
            }
            else
            {
                //Caso não exista, muda a cor do texto para vermelho
                caminhoPermTextBox.Foreground = Brushes.Red;
                caminhoPermTextBox.BorderBrush = Brushes.Red;

                //Exibe mensagem que o caminho não existe
                MessageBox.Show("Caminho dos arquivos DSS não existe", "Erro!");
            }
        }

        // seleciona caminho path recursos permanentes
        private void caminhoPermBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                caminhoPermTextBox.Text = dialog.SelectedPath;
                caminhoDSSTB_LostFocus(caminhoPermTextBox, new RoutedEventArgs());
            }
        }

        // Adiciona PU
        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            _parGUI._tensaoSaidaBarUsuario = tensaoSaidaBarTextBox.Text;

            double tensaoPU = Double.Parse(_parGUI._tensaoSaidaBarUsuario);
            
            tensaoPU = tensaoPU + 0.01;

            _parGUI._tensaoSaidaBarUsuario = tensaoPU.ToString();

            tensaoSaidaBarTextBox.Text = tensaoPU.ToString();
        }

        // habilita Text, caso conformo 
        private void usarTensoesBarramentoCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            // desabilita TextBox valor default pu
            tensaoSaidaBarTextBox.IsEnabled = false;
        }

        private void usarTensoesBarramentoCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            // habilita TextBox valor default pu
            tensaoSaidaBarTextBox.IsEnabled = true;
        }

        private void calculaPUotm_Checked(object sender, RoutedEventArgs e)
        {
            //desabilita PUSaidaSE arquivo
            usarTensoesBarramentoCheckBox.IsChecked = false;
        }

        private void horaTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            double horaInt = Int32.Parse(horaTextBox.Text);

            // verifica se hora eh valida
            if (horaInt < 0 || horaInt > 23)
            {
                //TODO
                //UpdateDisplay("Entre com uma hora válida!");
                Disp("Entre com uma hora válida!");
            }
            else
            {
                //TODO nao dar update no display
                _parGUI._hora = horaTextBox.Text;
                //UpdateDisplay("Nova hora definida");
                Disp("Nova hora definida");
            }
        }

        private void OpcoesAvancadasButton_Click(object sender, RoutedEventArgs e)
        {
            // preenche variaveis da GUI antes da exeucao
            _parGUI.preencheVariaveis(this);

            // janela OpcoesAvancadas
            OpcoesAvancadas opcoes = new OpcoesAvancadas(this, _parGUI, _parAvan);

            //
            if (opcoes.ShowDialog() == true)
            {
                _parAvan = opcoes._parAvan;
            }
        }
    }
}
