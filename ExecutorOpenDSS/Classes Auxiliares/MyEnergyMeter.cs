using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes_Auxiliares
{
    class MyEnergyMeter
    {
        public double MaxkW = 0;
        public double MaxkWLosses = 0;
        public double kWh = 0;
        public double kvarh = 0;
        public double LossesKWh = 0;
        public double TransformerLosses = 0;
        public double MTLineLosses = 0;
        public double BTLineLosses = 0;
        public double lineLossesPosMode = 0;
        public double lineLossesZeroMode = 0;
        public double NoLoadLosseskWh = 0;
        public double MTEnergy = 0;
        public double BTEnergy = 0;
        public string sMes = "0";

        /*
        Reg 33 = 34.5 kV Losses
        Reg 34 = 13.8 kV Losses
        Reg 35 = 0.24 kV Losses
        Reg 36 = 0.22 kV Losses
         * */

        public MyEnergyMeter(MyEnergyMeter em)
        {
            this.kWh = em.kWh;
            this.kvarh = em.kvarh;
            this.LossesKWh = em.LossesKWh;
            this.TransformerLosses = em.TransformerLosses;
            this.MTLineLosses = em.MTLineLosses;
            this.BTLineLosses = em.BTLineLosses;
            this.lineLossesPosMode = em.lineLossesPosMode;
            this.lineLossesZeroMode = em.lineLossesZeroMode;
            this.NoLoadLosseskWh = em.NoLoadLosseskWh;
            this.MTEnergy = em.MTEnergy;
            this.BTEnergy = em.BTEnergy;
            this.sMes = em.sMes;
        }

        public MyEnergyMeter()
        {
            // TODO: Complete member initialization
        }

        public void geracaoPerdasPotEPerdasEnergiaType(List<double> _geracaoPerdasPotEPerdasEnergia)
        {
            MaxkW = _geracaoPerdasPotEPerdasEnergia[0];
            MaxkWLosses = _geracaoPerdasPotEPerdasEnergia[1];
            kWh = _geracaoPerdasPotEPerdasEnergia[2];
            kvarh = _geracaoPerdasPotEPerdasEnergia[3];
            LossesKWh = _geracaoPerdasPotEPerdasEnergia[4];
            TransformerLosses = _geracaoPerdasPotEPerdasEnergia[5];
            MTLineLosses = _geracaoPerdasPotEPerdasEnergia[6];
            BTLineLosses = _geracaoPerdasPotEPerdasEnergia[7];
            lineLossesPosMode = _geracaoPerdasPotEPerdasEnergia[8];
            lineLossesZeroMode = _geracaoPerdasPotEPerdasEnergia[9];
            NoLoadLosseskWh = _geracaoPerdasPotEPerdasEnergia[10];
            MTEnergy = _geracaoPerdasPotEPerdasEnergia[11];
            BTEnergy = _geracaoPerdasPotEPerdasEnergia[12];
        }

        /* //OLD CODE
        //TODO fazer operador * 
        //Soma operadores de energia 
        public static MyEnergyMeter operator +(MyEnergyMeter em1, MyEnergyMeter em2)
        {
            MyEnergyMeter ret = new MyEnergyMeter(em1);

            ret.kWh += em2.kWh;
            ret.kvarh += em2.kvarh;
            ret.LossesKWh += em2.LossesKWh;
            ret.TransformerLosses += em2.TransformerLosses;
            ret.MTLineLosses += em2.MTLineLosses;
            ret.BTLineLosses += em2.BTLineLosses;
            ret.lineLossesPosMode += em2.lineLossesPosMode;
            ret.lineLossesZeroMode += em2.lineLossesZeroMode;
            ret.NoLoadLosseskWh += em2.NoLoadLosseskWh;
            ret.MTEnergy += em2.MTEnergy;
            ret.BTEnergy += em2.BTEnergy;

            return ret;
        }*/

        //TODO fazer operador * 
        //Soma operadores de energia 
        public void Soma(MyEnergyMeter em)
        {
            this.kWh += em.kWh;
            this.kvarh += em.kvarh;
            this.LossesKWh += em.LossesKWh;
            this.TransformerLosses += em.TransformerLosses;
            this.MTLineLosses += em.MTLineLosses;
            this.BTLineLosses += em.BTLineLosses;
            this.lineLossesPosMode += em.lineLossesPosMode;
            this.lineLossesZeroMode += em.lineLossesZeroMode;
            this.NoLoadLosseskWh += em.NoLoadLosseskWh;
            this.MTEnergy += em.MTEnergy;
            this.BTEnergy += em.BTEnergy;
        }

        internal void AddEnergia(List<double> injecaoEPerdasEnergia)
        {
            kWh = injecaoEPerdasEnergia[0];
            kvarh = injecaoEPerdasEnergia[1];
            LossesKWh = injecaoEPerdasEnergia[2];
            TransformerLosses = injecaoEPerdasEnergia[3];
            MTLineLosses = injecaoEPerdasEnergia[4];
            BTLineLosses = injecaoEPerdasEnergia[5];
            lineLossesPosMode = injecaoEPerdasEnergia[6];
            lineLossesZeroMode = injecaoEPerdasEnergia[7];
            NoLoadLosseskWh = injecaoEPerdasEnergia[8];
            MTEnergy = injecaoEPerdasEnergia[9];
            BTEnergy = injecaoEPerdasEnergia[10];
        }

        public string formataResultado(string nomeAlim)
        {
            string conteudo;
            conteudo = nomeAlim + "\t";

            //Obtem medidores
            conteudo = conteudo + MaxkW.ToString("0.0000") + "\t"; //MaxkW
            conteudo = conteudo + MaxkWLosses.ToString("0.0000") + "\t"; //MaxkWLosses

            //Obtem medidores
            conteudo = conteudo + kWh.ToString("0.0000") + "\t"; //kW
            conteudo = conteudo + kvarh.ToString("0.0000") + "\t"; //kvarh
            conteudo = conteudo + LossesKWh.ToString("0.0000") + "\t"; //LossesKWh

            //perdas em transformadores
            conteudo = conteudo + TransformerLosses.ToString("0.0000") + "\t"; //TransformerLosses

            //perdas MT e BT
            conteudo = conteudo + MTLineLosses.ToString("0.0000") + "\t"; //MTLosses
            conteudo = conteudo + BTLineLosses.ToString("0.0000") + "\t"; //BTLosses

            //Line losses. Struct com os seguintes campos:
            conteudo = conteudo + lineLossesPosMode.ToString("0.0000") + "\t"; //lineLossesPosMode
            conteudo = conteudo + lineLossesZeroMode.ToString("0.0000") + "\t"; //lineLossesZeroMode

            // NoLoadLosses
            conteudo = conteudo + NoLoadLosseskWh.ToString("0.0000") + "\t"; //NoLoadLosses

            // MT Energy
            conteudo = conteudo + MTEnergy.ToString("0.0000") + "\t"; //

            // BT Energy
            conteudo = conteudo + BTEnergy.ToString("0.0000") + "\t"; //

            //mes
            conteudo = conteudo + sMes + "\t"; //

            return conteudo;
        }
        
        // seta o mes do calculo
        public void SetMes(int iMes, bool modoMensal)
        {
            //TODO FIX ME
            if (modoMensal)
            {
                sMes = iMes.ToString();
            }
        }
    }
}
