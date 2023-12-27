﻿using ExecutorOpenDSS.AuxClasses;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExecutorOpenDSS.MainClasses
{
    class LoopAnalysis
    {
        private readonly GeneralParameters _paramGerais;
        private DailyFlow _fluxoSoMT;

        public LoopAnalysis(GeneralParameters paramGerais)
        {
            _paramGerais = paramGerais;

            //Lê os alimentadores e armazena a lista de alimentadores 
            List<string> alimentadores = CemigFeeders.GetTodos(_paramGerais._parGUI.GetArqLstAlimentadores());

            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            // analisa cada alimentador
            foreach (string nomeAlim in alimentadores)
            {
                AnaliseLoopsPvt(nomeAlim);
            }

            // Grava Log // TODO
            //paramGerais._mWindow.GravaLog();
        }

        private void AnaliseLoopsPvt(string nomeAlim)
        {
            // atribui nomeAlim
            _paramGerais.SetNomeAlimAtual(nomeAlim);

            // Carrega arquivos DSS so MT
            _fluxoSoMT = new DailyFlow(_paramGerais, "DU", true);

            // TODO testar
            bool ret = _fluxoSoMT.LoadStringListwithDSSCommands();

            if (ret)
            {
                ret = _fluxoSoMT.ExecutaFluxoSnap();
            }

            // SE executou fluxo snap
            if (ret)
            {
                // verifica cancelamento usuario 
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return;
                }

                //
                AnaliseLoops2();
            }

        }

        // Analise de Loops
        private void AnaliseLoops2()
        {
            // Obtem loops
            string[] loops = _fluxoSoMT._oDSS.GetActiveCircuit().Topology.AllLoopedPairs;

            // armazena loops em lista
            List<string> lstLoops = new List<string>(loops);

            //Plota Looped Pairs 
            PlotaLoopedPairs(lstLoops);
        }

        //Plota niveis tensao nas barras dos trafos
        public void PlotaLoopedPairs(List<string> lstLoops)
        {
            // nome alim
            string nomeAlim = _paramGerais.GetNomeAlimAtual();

            // linha 
            String linha = "";

            // para cada key value
            foreach (string loop in lstLoops)
            {
                //armazena  nomeALim e loop
                linha += nomeAlim + "\t" + loop;

                //adiciona quebra de linha
                if (loop != lstLoops.Last())
                {
                    linha += "\n";
                }
            }
            TxtFile.GravaEmArquivo(linha, _paramGerais.GetNomeCompArqLoops(), _paramGerais._mWindow);
        }
    }
}
