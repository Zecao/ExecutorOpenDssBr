namespace ExecutorOpenDSS.AuxClasses
{
    class MyEnergyMeter
    {
        public double MaxkW = 0;
        public double MaxkWLosses = 0;
        public double KWh = 0;
        public double kVAr_h = 0;
        public double LossesKWh = 0;
        public double TransformerLosses = 0;
        public double BTLineLosses = 0;
        public double lineLossesPosMode = 0;
        public double lineLossesZeroMode = 0;
        public double NoLoadLosseskWh = 0;
        public double MTEnergy = 0;
        public double BTEnergy = 0;
        public string _sMes = "0";
        public double MTLineLosses34KV = 0;
        public double MTLineLosses = 0;
        public double TransformerAllLosses34KV = 0;
        public double kWh_GD = 0;
        public double kVArh_GD = 0;
        public double loadMultAlim = 1;

        // Construtor por copia
        public MyEnergyMeter(MyEnergyMeter em)
        {
            KWh = em.KWh;
            kVAr_h = em.kVAr_h;
            LossesKWh = em.LossesKWh;
            TransformerLosses = em.TransformerLosses;
            MTLineLosses = em.MTLineLosses;
            BTLineLosses = em.BTLineLosses;
            lineLossesPosMode = em.lineLossesPosMode;
            lineLossesZeroMode = em.lineLossesZeroMode;
            NoLoadLosseskWh = em.NoLoadLosseskWh;
            MTEnergy = em.MTEnergy;
            BTEnergy = em.BTEnergy;
            _sMes = em._sMes;
            MTLineLosses34KV = em.MTLineLosses34KV;
            TransformerAllLosses34KV = em.TransformerAllLosses34KV;
            kWh_GD = em.kWh_GD;
            kVArh_GD = em.kVArh_GD;
            loadMultAlim = em.loadMultAlim;
        }

        public void GravaLoadMult(double LM)
        {
            loadMultAlim = LM;
        }

        public MyEnergyMeter()
        {
        }

        // operador soma  de energia 
        public void Soma(MyEnergyMeter em)
        {
            KWh += em.KWh;
            kVAr_h += em.kVAr_h;
            LossesKWh += em.LossesKWh;
            TransformerLosses += em.TransformerLosses;
            MTLineLosses += em.MTLineLosses;
            BTLineLosses += em.BTLineLosses;
            lineLossesPosMode += em.lineLossesPosMode;
            lineLossesZeroMode += em.lineLossesZeroMode;
            NoLoadLosseskWh += em.NoLoadLosseskWh;
            MTEnergy += em.MTEnergy;
            BTEnergy += em.BTEnergy;
            MTLineLosses34KV += em.MTLineLosses34KV;
            TransformerAllLosses34KV += em.TransformerAllLosses34KV;
            kWh_GD += em.kWh_GD;
            kVArh_GD += em.kVArh_GD;
        }

        internal void MultiplicaEnergia(int numDias)
        {
            KWh *= numDias;
            kVAr_h *= numDias;
            LossesKWh *= numDias;
            TransformerLosses *= numDias;
            MTLineLosses *= numDias;
            BTLineLosses *= numDias;
            lineLossesPosMode *= numDias;
            lineLossesZeroMode *= numDias;
            NoLoadLosseskWh *= numDias;
            MTEnergy *= numDias;
            BTEnergy *= numDias;
            MTLineLosses34KV *= numDias;
            TransformerAllLosses34KV *= numDias;
            kWh_GD *= numDias;
            kVArh_GD *= numDias;
        }

        public string FormataResultado(string nomeAlim)
        {
            string conteudo = "";
            conteudo += nomeAlim + "\t";

            //Obtem medidores
            conteudo += MaxkW.ToString("0.0000") + "\t"; //MaxkW //2
            conteudo += MaxkWLosses.ToString("0.0000") + "\t"; //MaxkWLosses //3

            //Obtem medidores
            conteudo += KWh.ToString("0.0000") + "\t"; //kW //4
            conteudo += kVAr_h.ToString("0.0000") + "\t"; //kvarh //5
            conteudo += LossesKWh.ToString("0.0000") + "\t"; //LossesKWh //6

            //perdas em transformadores
            conteudo += TransformerLosses.ToString("0.0000") + "\t"; //TransformerLosses //7

            //perdas MT e BT
            conteudo += MTLineLosses.ToString("0.0000") + "\t"; //MTLosses //8
            conteudo += BTLineLosses.ToString("0.0000") + "\t"; //BTLosses //9
            /*
            //Line losses. Struct com os seguintes campos:
            conteudo += lineLossesPosMode.ToString("0.0000") + "\t"; //lineLossesPosMode
            conteudo += lineLossesZeroMode.ToString("0.0000") + "\t"; //lineLossesZeroMode
            */
            conteudo += kWh_GD.ToString("0.0000") + "\t"; //kW_GD //10
            conteudo += kVArh_GD.ToString("0.0000") + "\t"; //kVAr_GD

            // NoLoadLosses
            conteudo += NoLoadLosseskWh.ToString("0.0000") + "\t"; //NoLoadLosses

            // MT Energy
            conteudo += MTEnergy.ToString("0.0000") + "\t"; //

            // BT Energy
            conteudo += BTEnergy.ToString("0.0000") + "\t"; //

            //  MTLineLosses34KV
            conteudo += MTLineLosses34KV.ToString("0.0000") + "\t"; //

            // TransformerAllLosses34KV
            conteudo += TransformerAllLosses34KV.ToString("0.0000") + "\t"; //

            // mes
            conteudo += _sMes + "\t"; //

            //load mult
            conteudo += loadMultAlim.ToString("0.0000") + "\t";

            return conteudo;
        }

        // seta o mes do calculo
        public void SetMesEM(int iMes)
        {
            _sMes = iMes.ToString();
        }
    }
}
