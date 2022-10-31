
/*� ���, ��� ����������, ��������� ����� ������ ������ � ������ ��������� ������ ���������
 */

using Excel = Microsoft.Office.Interop.Excel; // ����������� ������ ��� ��� ������ � ������


namespace AHU_Configurator
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }
        double ComAmps = 0; // ����� ������� ���
        double Amps = 0;
        double MotorAmps = 0;
        double AmpsSup = 0;
        double AmpsExh = 0;
        double AmpsEl = 0;
        double AmpsElTerminal = 0;
        double AmpsMotor = 0;

        string motorProtector = "";
        string motorPtotectorSup = "";
        string motorPtotectorExh = "";

        int pieceOfMotorProtector = 0;
        int pieceOfMotorPtotectorExh = 0;
        int pieceOfMotorPtotectorSup = 0;

        string mainSwitch = "";
        int piceOfMainSwitch = 1;
        int piceOfHandleSwitch = 1;
        string cabinet = "";
        int piceOfCabinet = 1;
        string AutomaticSwitchOnePhase = "";
        int piceOfAutomaticSwitchOnePhase = 1; // ������������ ��������
        string AutomaticSwitchThreePhase = "";
        int piceOfAutomaticSwitchThreePhase = 0; // ������������ ��������
        int piceOfRelay = 3; // ���������� ����
        int piceSocketOfrelay = 0; // ���������� ������� ��� ����

        int blackTerminal2_5 = 0;
        int blueTerminal2_5 = 0;
        int gndTerminal2_5 = 0;
        int ElblackTerminal2_5 = 0;
        int ElblueTerminal2_5 = 0;
        int ElgndTerminal2_5 = 0;
        int ElblackTerminal4 = 0;
        int ElgndTerminal4 = 0;
        int ElblackTerminal6 = 0;
        int ElgndTerminal6 = 0;
        int ElblackTerminal10 = 0;
        int ElgndTerminal10 = 0;
        int ElblackTerminal16 = 0;
        int ElgndTerminal16 = 0;

        int MotorblackTerminal2_5 = 0;
        int MotorblueTerminal2_5 = 0;
        int MotorgndTerminal2_5 = 0;
        int MotorblackTerminal4 = 0;
        int MotorgndTerminal4 = 0;
        int MotorblackTerminal6 = 0;
        int MotorgndTerminal6 = 0;
        int MotorblackTerminal10 = 0;
        int MotorgndTerminal10 = 0;
        int MotorblackTerminal16 = 0;
        int MotorgndTerminal16 = 0;

        int ComblackTerminal2_5 = 0;
        int ComblueTerminal2_5 = 0;
        int ComgndTerminal2_5 = 0;
        int ComblackTerminal4 = 0;
        int ComgndTerminal4 = 0;
        int ComblackTerminal6 = 0;
        int ComgndTerminal6 = 0;
        int ComblackTerminal10 = 0;
        int ComgndTerminal10 = 0;
        int ComblackTerminal16 = 0;
        int ComgndTerminal16 = 0;

        int MotorblueTerminal4 = 0;
        int MotorblueTerminal6 = 0;
        int MotorblueTerminal10 = 0;
        int MotorblueTerminal16 = 0;

        int ElblueTerminal4 = 0;
        int ElblueTerminal6 = 0;
        int ElblueTerminal10 = 0;
        int ElblueTerminal16 = 0;

        int NumOfDO = 0;
        int NumOfAO = 0;
        int NumOfDI = 0;
        int NumOfAI = 0;
        int NumOfAirValSup = 0;
        int NumOfAirValExh = 0;
        int NumOfDO_El = 0; 
        int NumOfDO_Dry = 0;
        int NumOfDO_Hum = 0;
        int NumOfDO_Cold = 0;
        int NumOfDO_AirValve = 0;
        int NumOfDO_Motor = 0; 
        int NumOfDO_MotorSup = 0;
        int NumOfDO_MotorExh = 0; 


        int relay2pk = 1;
        int SocketRelay2pk = 1;
        string NameRelay2pk = "�� slim 22/2 5A 230� AC EKF AVERES �������: rps-22-2-230";
        string NameSocketRelay2pk = "�M slim 22/2 EKF AVERES �������: rms-22-2";

        int CircuitBreakerElThreePhase = 0;
        int CircuitBreakerPumpThreePhase = 0;
        int contactorEl = 0;
        int contacorSup = 0;
        int contactorExh = 0;
        string NameCircuitBreakerElThreePhas = "";
        string NameCircuitBreakerPumpThreePhase = "";
        string NameContactor = "";
        string NameContactorEl = "";
        string NameContactorExh = "";


        //    string NameBlackTerminal2_5 = "";
        //    string NameBlueTerminal2_5 = "";
        //    string NameGndTerminal2_5 = "";
        //    string NameBlackTerminal4 = "";
        //    string NameGndTerminal4 = "";
        //    string NameBlackTerminal6 = "";
        //    string NameGndTerminal6 = "";
        //    string NameBlackTerminal10 = "";
        //    string NameGndTerminal10 = "";
        //    string NameBlackTerminal16 = "";
        //    string NameGndTerminal16 = "";

        string NameBlackTerminal2_5 = "������� �������� JXB-2.5/35 ����� EKF PROxima �������: plc-jxb-2.4/35gy  ";
        string NameBlueTerminal2_5 = "������� �������� JXB-2.5/35 ����� EKF PROxima �������: plc - jxb - 2.5 / 35b";
        string NameGndTerminal2_5 = "������ �������� ��-JXB-2,5 ��� ���������� EKF �������: plc-ek-2.5/25";

        string NameBlackTerminal4 = "������� �������� JXB-4/35 ����� EKF PROxima �������: plc - jxb - 4 / 35gy";
        string NameBlueTerminal4 = "������� �������� JXB-4/35 ����� EKF PROxima �������: plc-jxb-4/35b";
        string NameGndTerminal4 = "������ �������� ��-JXB-4 ��� ���������� EKF �������: plc - ek - 4 / 32";

        string NameBlackTerminal6 = "������� �������� JXB-6/35 ����� EKF PROxima �������: plc - jxb - 6 / 35gy";
        string NameBlueTerminal6 = "������� �������� JXB-6/35 ����� EKF PROxima �������: plc-jxb-6/35b";
        string NameGndTerminal6 = "������ �������� ��-JXB-6 ��� ���������� EKF �������: plc - ek - 6 / 40";

        string NameBlackTerminal10 = "������� �������� JXB-10/35 ����� EKF PROxima �������: plc - jxb - 10 / 35gy";
        string NameBlueTerminal10 = "������� �������� JXB-10/35 ����� EKF PROxima �������: plc-jxb-10/35b";
        string NameGndTerminal10 = "������ �������� ��-JXB-10 ��� ���������� EKF �������: plc - ek - 10 / 63";

        string NameBlackTerminal16 = "������� �������� JXB-16/35 ����� EKF PROxima �������: plc - jxb - 16 / 35gy";
        string NameBlueTerminal16 = "������� �������� JXB-16/35 ����� EKF PROxima �������: plc-jxb-16/35b";
        string NameGndTerminal16 = "������ �������� ��-JXB-16 ��� ���������� EKF �������: plc - ek - 16 / 80";



        private void button1_Click(object sender, EventArgs e)
        {



            // ��������� ��������
            if (checkBoxSupply.Checked & !checkBoxExh.Checked)

            {
                /* ����� ���������� ��� ������ ����������� ���� */
                if (ThreePhaseSup.Checked)
                {
                    double pwr = Double.Parse(PoweSupMain.Text);
                    if ((Reserve.Checked) ^ (duobleReserve.Checked))
                    {
                        ComAmps = ((pwr / 380) * 2)+(AmpsEl);
                        AmpsSup = pwr / 380;
                        //  test.Text = Amps.ToString();
                        //  test1.Text = ComAmps.ToString();
                    }
                    else
                    {
                        ComAmps = (pwr / 380) + AmpsEl;
                        AmpsSup = ComAmps;
                        // test.Text = ComAmps.ToString();
                        //  test1.Text = ComAmps.ToString();
                    }
                }
                if (OneNumPhase.Checked)
                {
                    double pwr = Double.Parse(PoweSupMain.Text);
                    if ((Reserve.Checked) ^ (duobleReserve.Checked))
                    {
                        ComAmps = ((pwr / 220) * 2) + AmpsEl;
                        AmpsSup = pwr / 220;
                        //test.Text = ComAmps.ToString();
                        // test1.Text = ComAmps.ToString();
                    }
                    else
                        ComAmps = (pwr / 220) + AmpsEl;
                    AmpsSup = ComAmps;
                    //  test.Text = Amps.ToString();
                    //  test1.Text = ComAmps.ToString();
                }

                // ����� �����
                { /* ����� ����� */
                    if (ComAmps < 83)
                    { cabinet = "������ ������� ST ��� �/� 800x600x250 �������: R5ST0869WMP "; }
                    else if (ComAmps < 200)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1000 x 600 x 250 �� (� � � � �) �������: R5ST1069"; }
                    else if (ComAmps < 400)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1200 x 800 x 300 �� (� � � � �) �������: R5ST1283"; }
                    else if (ComAmps < 630)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1400 x 800 x 300 �� (� � � � �) �������: R5ST1483"; }
                    else if (ComAmps < 1000)
                    { cabinet = "DKC ��� ��������� ��� 1600�800�400 IP31 ���� ��� ��������� ������ ������������� ���-16.8.4-0 �������: YKM40-1684-31"; }
                }


                /* ����� ���������� */
                {
                    if (ComAmps < 40)
                    { mainSwitch = "��������� 40A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 40 - 3p - f"; }
                    else if (ComAmps < 63)
                    { mainSwitch = "��������� 63A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 63 - 3p - f"; }
                    else if (ComAmps < 83)
                    { mainSwitch = "��������� 80A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF PROxima �������: tb - 80 - 3p - f"; }
                    else if (ComAmps < 160)
                    { mainSwitch = "��������� 160A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 160 - 3p"; }
                    else if (ComAmps < 200)
                    { mainSwitch = "��������� 200A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 200 - 3p"; }
                    else if (ComAmps < 250)
                    { mainSwitch = "��������� 250A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 250 - 3p"; }
                    else if (ComAmps < 315)
                    { mainSwitch = "��������� 315A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 315 - 3p"; }
                    else if (ComAmps < 400)
                    { mainSwitch = "��������� 400A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 400 - 3p"; }
                    else if (ComAmps < 630)
                    { mainSwitch = "��������� 630A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 630 - 3p"; }
                    else if (ComAmps < 800)
                    { mainSwitch = "��������� 800A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 800 - 3p"; }
                    else if (ComAmps < 1000)
                    { mainSwitch = "��������� 1000A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 1000 - 3p"; }
                    if (ComAmps > 200)
                    {
                        piceOfHandleSwitch = piceOfMainSwitch;
                    }
                }

                /* ����� �������� ������ ���������*/

                {
                    // ���� ������ ������� ������
                    if (ThreePhaseSup.Checked)
                    {
                        if (AmpsSup < 0.63)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                        }
                        else if (AmpsSup < 1.0)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                        }
                        else if (AmpsSup < 1.2)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                        }
                        else if (AmpsSup < 2.2)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                        }
                        else if (AmpsSup < 3.6)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                        }
                        else if (AmpsSup < 5.6)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                        }
                        else if (AmpsSup < 9)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                        }
                        else if (AmpsSup < 13.0)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                        }
                        else if (AmpsSup < 17.0)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsSup < 22)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 17-23 � EKF PROxima �������: gv2p21 - pro";
                        }
                        else if (AmpsSup < 24.0)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 20-25 � EKF PROxima �������: gv2p22 - pro";
                        }
                        else if (AmpsSup < 31.0)
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 24-32 � EKF PROxima �������: gv2p32 - pro";
                        }

                        // ���������� ��������� ������ ����������

                        if ((Reserve.Checked) || (duobleReserve.Checked))
                        {
                            if (AmpsSup < 30)
                            {
                                MotorblackTerminal2_5 = 6;
                                MotorgndTerminal2_5 = 2;
                            }
                            else if (AmpsSup < 41)
                            {
                                MotorblackTerminal4 = 6;
                                MotorgndTerminal4 = 2;
                            }
                            else if (AmpsSup < 50)
                            {
                                MotorblackTerminal6 = 6;
                                MotorgndTerminal6 = 2;
                            }
                            else if (AmpsSup < 70)
                            {
                                MotorblackTerminal10 = 6;
                                MotorgndTerminal10 = 2;
                            }
                            else if (AmpsSup < 100)
                            {
                                MotorblackTerminal16 = 6;
                                MotorgndTerminal16 = 2;
                            }
                            pieceOfMotorPtotectorSup = 2;
                        }

                        else if (WithOutReserv.Checked)
                        {
                            if (AmpsSup < 30)
                            {
                                MotorblackTerminal2_5 = 3;
                                MotorgndTerminal2_5 = 1;
                            }
                            else if (AmpsSup < 41)
                            {
                                MotorblackTerminal4 = 3;
                                MotorgndTerminal4 = 1;
                            }
                            else if (AmpsSup < 50)
                            {
                                MotorblackTerminal6 = 3;
                                MotorgndTerminal6 = 1;
                            }
                            else if (AmpsSup < 70)
                            {
                                MotorblackTerminal10 = 3;
                                MotorgndTerminal10 = 1;
                            }
                            else if (AmpsSup < 100)
                            {
                                MotorblackTerminal16 = 3;
                                MotorgndTerminal16 = 1;
                            }
                            pieceOfMotorPtotectorSup = 1;
                        }
                        // ���������� ���������� ��� ����������
                        if (AmpsSup < 9)
                        {
                            NameContactor = "��������� ��� 9� 1NO 230� �� EKF AVERES �������: ctr-s-9-10-230-av";

                        }
                        else if (AmpsSup < 12)
                        {
                            NameContactor = "��������� ��� 12� 1NO 230� �� EKF AVERES �������: ctr-s-12-10-230-av";
                        }
                        else if (AmpsSup < 18)
                        {
                            NameContactor = "��������� ��� 18� 1NO 230� �� EKF AVERES �������: ctr-s-18-10-230-av";
                        }
                        else if (AmpsSup < 22)
                        {
                            NameContactor = "��������� ��� 22� 1NO 230� �� EKF AVERES �������: ctr-s-22-10-230-av";
                        }
                        else if (AmpsSup < 25)
                        {
                            NameContactor = "��������� ��� 25� 230� �� EKF AVERES �������: ctr-s-25-00-230-av";
                        }
                        else if (AmpsSup < 30)
                        {
                            NameContactor = "��������� ��� 30� 230� �� EKF AVERES �������: ctr-s-30-00-230-av";
                        }
                        else if (AmpsSup < 32)
                        {
                            NameContactor = "��������� ��� 32� 230� �� EKF AVERES �������: ctr-s-32-00-230-av";
                        }
                        else if (AmpsSup < 38)
                        {
                            NameContactor = "��������� ��� 38� 230� �� EKF AVERES �������: ctr-s-40-00-230-av";
                        }
                        else if (AmpsSup < 50)
                        {
                            NameContactor = "��������� ��� 50� 230� �� EKF AVERES �������: ctr-s-50-00-230-av";
                        }
                        else if (AmpsSup < 60)
                        {
                            NameContactor = "��������� ��� 60� 230� �� EKF AVERES �������: ctr-s-60-00-230-av";
                        }
                        else if (AmpsSup < 65)
                        {
                            NameContactor = "��������� ��� 65� 230� �� EKF AVERES �������: ctr-s-70-00-230-av";
                        }
                        else if (AmpsSup < 80)
                        {
                            NameContactor = "��������� ��� 80� 230� �� EKF AVERES �������: ctr-s-80-00-230-av";
                        }
                        else if (AmpsSup < 90)
                        {
                            NameContactor = "��������� ��� 90� 230� �� EKF AVERES �������: ctr-s-90-00-230-av";
                        }
                        else if (AmpsSup < 100)
                        {
                            NameContactor = "��������� ��� 100� 230� �� EKF AVERES �������: ctr-s-100-00-230-av";
                        }
                    }

                    // ���������� �������
                    if (OneNumPhase.Checked)
                    {
                        if (AmpsSup < 1.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 1A (D) 10kA EKF AVERES �������: mcb10 - 1 - 01D - av";
                        }
                        else if (AmpsSup < 2.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 2A (D) 10kA EKF AVERES �������: mcb10 - 1 - 02D - av";
                        }
                        else if (AmpsSup < 4.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 4A (D) 10kA EKF AVERES �������: mcb10 - 1 - 04D - av";
                        }
                        else if (AmpsSup < 6.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 6A (D) 10kA EKF AVERES �������: mcb10 - 1 - 06D - av";
                        }
                        else if (AmpsSup < 10.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 6A (D) 10kA EKF AVERES �������: mcb10 - 1 - 06D - av";
                        }
                        else if (AmpsSup < 16.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 16A (D) 10kA EKF AVERES �������: mcb10 - 1 - 16D - av";
                        }
                        else if (AmpsSup < 20.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 20A (D) 10kA EKF AVERES �������: mcb10 - 1 - 20D - av";
                        }
                        else if (AmpsSup < 25.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 25A (D) 10kA EKF AVERES �������: mcb10 - 1 - 25D - av";
                        }
                        else if (AmpsSup < 32.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 32A (D) 10kA EKF AVERES �������: mcb10 - 1 - 32D - av";
                        }
                        else if (AmpsSup < 40)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 40A (D) 10kA EKF AVERES �������: mcb10 - 1 - 40D - av";
                        }
                        else if (AmpsSup < 50.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 50A (D) 10kA EKF AVERES �������: mcb10 - 1 - 50D - av";
                        }
                        else if (AmpsSup < 63.0)
                        {
                            motorPtotectorSup = "����������� �������������� AV-10 1P 63A (D) 10kA EKF AVERES �������: mcb10 - 1 - 63D - av";
                        }

                        // ���������� ��������� ������ ���������� � �����

                        if ((Reserve.Checked) || (duobleReserve.Checked))
                        {
                            // ���������� ����� ��� ���������� ����
                            if (AmpsSup < 30)
                            {
                                MotorblackTerminal2_5 = 2;
                                MotorgndTerminal2_5 = 2;
                                MotorblueTerminal2_5 = 2;
                            }
                            else if (AmpsSup < 41)
                            {
                                MotorblackTerminal4 = 2;
                                MotorgndTerminal4 = 2;
                                MotorblueTerminal4 = 2;
                            }
                            else if (AmpsSup < 50)
                            {
                                MotorblackTerminal6 = 2;
                                MotorgndTerminal6 = 2;
                                MotorblueTerminal6 = 2;
                            }
                            else if (AmpsSup < 70)
                            {
                                MotorblackTerminal10 = 2;
                                MotorgndTerminal10 = 2;
                                MotorblueTerminal10 = 2;
                            }
                            else if (AmpsSup < 100)
                            {
                                MotorblackTerminal16 = 2;
                                MotorgndTerminal16 = 2;
                                MotorblueTerminal16 = 2;

                            }
                            // ���������� ��������� ���������� ����
                            pieceOfMotorPtotectorSup = 2;

                        }

                        else if (WithOutReserv.Checked)
                        {
                            // ���������� ����� ��� ���������� ����
                            if (AmpsSup < 30)
                            {
                                MotorblackTerminal2_5 = 1;
                                MotorgndTerminal2_5 = 1;
                                MotorblueTerminal2_5 = 1;
                            }
                            else if (AmpsSup < 41)
                            {
                                MotorblackTerminal4 = 1;
                                MotorgndTerminal4 = 1;
                                MotorblueTerminal4 = 1;
                            }
                            else if (AmpsSup < 50)
                            {
                                MotorblackTerminal6 = 1;
                                MotorgndTerminal6 = 1;
                                MotorblueTerminal6 = 1;
                            }
                            else if (AmpsSup < 70)
                            {
                                MotorblackTerminal10 = 1;
                                MotorgndTerminal10 = 1;
                                MotorblueTerminal10 = 1;
                            }
                            else if (AmpsSup < 100)
                            {
                                MotorblackTerminal16 = 1;
                                MotorgndTerminal16 = 1;
                                MotorblueTerminal16 = 1;
                            }
                            // ���������� ��������� ���������� ����
                            pieceOfMotorPtotectorSup = 1;
                        }
                    }
                }

                /*����� ����������� ��������������� ����������� */
                {
                    AutomaticSwitchOnePhase = "����������� �������������� AV-6 1P 10A (C) 6kA EKF AVERES �������: mcb6 - 1 - 10C - av";
                    piceOfAutomaticSwitchOnePhase = piceOfAutomaticSwitchOnePhase + 1;
                    /* ����� ����������� ��������������� ����������� */
                }

                /* EXCEL � �������� �������� ����� */

                /* ������ � ������ */
                // ��������� ����� ������
                //��������� ����������

                Excel.Application app = new Excel.Application
                {
                    //���������� Excel
                    Visible = true,
                    //���������� ������ � ������� �����
                    SheetsInNewWorkbook = 2
                };

                //�������� ������� �����
                Excel.Workbook workBook = app.Workbooks.Add(Type.Missing);

                //��������� ����������� ���� � �����������
                app.DisplayAlerts = false;

                //�������� ������ ���� ��������� (���� ���������� � 1)
                Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);

                //��������� DO
              
             
                // �����������������
                if (checkBoxEl.Checked)
                {
                    // ����������
                    if (ThreePhaseEl.Checked) // ���������� �������� � ���������� ����� ����������
                    {
                        double pwrEl = Double.Parse(powerEl.Text);
                        AmpsEl = pwrEl / 380;
                        if (AmpsEl < 30)
                        {
                            int tempblackTerminal2_5 = ElblackTerminal2_5 + (((int)numericStepEl.Value) * 3);
                            ElblackTerminal2_5 = tempblackTerminal2_5;
                            ElgndTerminal2_5 = ElgndTerminal2_5 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;

                        }
                        else if (AmpsEl < 41)
                        {
                            int tempblackTerminal4 = ElblackTerminal4 + (((int)numericStepEl.Value) * 3);
                            ElblackTerminal4 = tempblackTerminal4;
                            ElgndTerminal4 = ElgndTerminal4 + 1;
                            NumOfDO_El = (int)numericStepEl.Value; ;


                        }
                        else if (AmpsEl < 50)
                        {
                            int tempblackTerminal6 = ElblackTerminal6 + (((int)numericStepEl.Value) * 3);
                            ElblackTerminal6 = tempblackTerminal6;
                            ElgndTerminal6 = ElgndTerminal6 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        else if (AmpsEl < 80)
                        {
                            int tempblackTerminal10 = ElblackTerminal10 + (((int)numericStepEl.Value) * 3);
                            ElblackTerminal10 = tempblackTerminal10;
                            ElgndTerminal10 = ElgndTerminal10 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        else if (AmpsEl < 100)
                        {
                            int tempblackTerminal16 = ElblackTerminal16 + (((int)numericStepEl.Value) * 3);
                            ElblackTerminal16 = tempblackTerminal16;
                            ElgndTerminal16 = ElgndTerminal16 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        if (checkBoxColdOnOff.Checked)
                        {
                            int tempblackTerminal2_5 = ElblackTerminal2_5 + ((int)numericStepCold.Value);
                            ElblackTerminal2_5 = tempblackTerminal2_5;
                            NumOfDO_El = (int)numericStepEl.Value;


                        }

                        if (AmpsEl < 9)
                        {
                            NameContactorEl = "��������� ��� 9� 1NO 230� �� EKF AVERES �������: ctr-s-9-10-230-av";

                        }
                        else if (AmpsEl < 12)
                        {
                            NameContactorEl = "��������� ��� 12� 1NO 230� �� EKF AVERES �������: ctr-s-12-10-230-av";
                        }
                        else if (AmpsEl < 18)
                        {
                            NameContactorEl = "��������� ��� 18� 1NO 230� �� EKF AVERES �������: ctr-s-18-10-230-av";
                        }
                        else if (AmpsEl < 22)
                        {
                            NameContactorEl = "��������� ��� 22� 1NO 230� �� EKF AVERES �������: ctr-s-22-10-230-av";
                        }
                        else if (AmpsEl < 25)
                        {
                            NameContactorEl = "��������� ��� 25� 230� �� EKF AVERES �������: ctr-s-25-00-230-av";
                        }
                        else if (AmpsEl < 30)
                        {
                            NameContactorEl = "��������� ��� 30� 230� �� EKF AVERES �������: ctr-s-30-00-230-av";
                        }
                        else if (AmpsEl < 32)
                        {
                            NameContactorEl = "��������� ��� 32� 230� �� EKF AVERES �������: ctr-s-32-00-230-av";
                        }
                        else if (AmpsEl < 38)
                        {
                            NameContactorEl = "��������� ��� 38� 230� �� EKF AVERES �������: ctr-s-40-00-230-av";
                        }
                        else if (AmpsEl < 50)
                        {
                            NameContactorEl = "��������� ��� 50� 230� �� EKF AVERES �������: ctr-s-50-00-230-av";
                        }
                        else if (AmpsEl < 60)
                        {
                            NameContactorEl = "��������� ��� 60� 230� �� EKF AVERES �������: ctr-s-60-00-230-av";
                        }
                        else if (AmpsEl < 65)
                        {
                            NameContactorEl = "��������� ��� 65� 230� �� EKF AVERES �������: ctr-s-70-00-230-av";
                        }
                        else if (AmpsEl < 80)
                        {
                            NameContactorEl = "��������� ��� 80� 230� �� EKF AVERES �������: ctr-s-80-00-230-av";
                        }
                        else if (AmpsEl < 90)
                        {
                            NameContactorEl = "��������� ��� 90� 230� �� EKF AVERES �������: ctr-s-90-00-230-av";
                        }
                        else if (AmpsEl < 100)
                        {
                            NameContactorEl = "��������� ��� 100� 230� �� EKF AVERES �������: ctr-s-100-00-230-av";
                        }
                        contactorEl = (int)numericStepEl.Value;

                    }
                    // ����������
                    else if (OnePhaseEl.Checked)
                    {
                        double pwrEl = Double.Parse(powerEl.Text);
                        AmpsEl = pwrEl / 220;
                        if (AmpsEl < 30)
                        {
                            int tempblackTerminal2_5 = ElblackTerminal2_5 + (((int)numericStepEl.Value));
                            ElblackTerminal2_5 = tempblackTerminal2_5;
                            ElgndTerminal2_5 = ElgndTerminal2_5 + 1;
                            ElblueTerminal2_5 = ElblueTerminal2_5 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        else if (AmpsEl < 41)
                        {
                            int tempblackTerminal4 = ElblackTerminal4 + (((int)numericStepEl.Value));
                            ElblackTerminal4 = tempblackTerminal4;
                            ElgndTerminal4 = ElgndTerminal4 + 1;
                            ElblueTerminal4 = ElblueTerminal4 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;

                        }
                        else if (AmpsEl < 50)
                        {
                            int tempblackTerminal6 = ElblackTerminal6 + (((int)numericStepEl.Value));
                            ElblackTerminal6 = tempblackTerminal6;
                            ElgndTerminal6 = ElgndTerminal6 + 1;
                            ElblueTerminal6 = ElblueTerminal6 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        else if (AmpsEl < 80)
                        {
                            int tempblackTerminal10 = ElblackTerminal10 + (((int)numericStepEl.Value));
                            ElblackTerminal10 = tempblackTerminal10;
                            ElgndTerminal10 = ElgndTerminal10 + 1;
                            ElblueTerminal10 = ElblueTerminal10 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        else if (AmpsEl < 100)
                        {
                            int tempblackTerminal16 = ElblackTerminal16 + (((int)numericStepEl.Value));
                            ElblackTerminal16 = tempblackTerminal16;
                            ElgndTerminal16 = ElgndTerminal16 + 1;
                            ElblueTerminal16 = ElblueTerminal16 + 1;
                            NumOfDO_El = (int)numericStepEl.Value;
                        }
                        if (AmpsEl < 9)
                        {
                            NameContactorEl = "��������� ��� 9� 1NO 230� �� EKF AVERES �������: ctr-s-9-10-230-av";

                        }
                        else if (AmpsEl < 12)
                        {
                            NameContactorEl = "��������� ��� 12� 1NO 230� �� EKF AVERES �������: ctr-s-12-10-230-av";
                        }
                        else if (AmpsEl < 18)
                        {
                            NameContactorEl = "��������� ��� 18� 1NO 230� �� EKF AVERES �������: ctr-s-18-10-230-av";
                        }
                        else if (AmpsEl < 22)
                        {
                            NameContactorEl = "��������� ��� 22� 1NO 230� �� EKF AVERES �������: ctr-s-22-10-230-av";
                        }
                        else if (AmpsEl < 25)
                        {
                            NameContactorEl = "��������� ��� 25� 230� �� EKF AVERES �������: ctr-s-25-00-230-av";
                        }
                        else if (AmpsEl < 30)
                        {
                            NameContactorEl = "��������� ��� 30� 230� �� EKF AVERES �������: ctr-s-30-00-230-av";
                        }
                        else if (AmpsEl < 32)
                        {
                            NameContactorEl = "��������� ��� 32� 230� �� EKF AVERES �������: ctr-s-32-00-230-av";
                        }
                        else if (AmpsEl < 38)
                        {
                            NameContactorEl = "��������� ��� 38� 230� �� EKF AVERES �������: ctr-s-40-00-230-av";
                        }
                        else if (AmpsEl < 50)
                        {
                            NameContactorEl = "��������� ��� 50� 230� �� EKF AVERES �������: ctr-s-50-00-230-av";
                        }
                        else if (AmpsEl < 60)
                        {
                            NameContactorEl = "��������� ��� 60� 230� �� EKF AVERES �������: ctr-s-60-00-230-av";
                        }
                        else if (AmpsEl < 65)
                        {
                            NameContactorEl = "��������� ��� 65� 230� �� EKF AVERES �������: ctr-s-70-00-230-av";
                        }
                        else if (AmpsEl < 80)
                        {
                            NameContactorEl = "��������� ��� 80� 230� �� EKF AVERES �������: ctr-s-80-00-230-av";
                        }
                        else if (AmpsEl < 90)
                        {
                            NameContactorEl = "��������� ��� 90� 230� �� EKF AVERES �������: ctr-s-90-00-230-av";
                        }
                        else if (AmpsEl < 100)
                        {
                            NameContactorEl = "��������� ��� 100� 230� �� EKF AVERES �������: ctr-s-100-00-230-av";
                        }
                        contactorEl = (int)numericStepEl.Value;

                    }
                }

                // ��������� ���������� ON-Off
                if (checkBoxColdOnOff.Checked)
                {
                    blackTerminal2_5 = (blackTerminal2_5 + (((int)numericStepCold.Value)) * 2);
                    NumOfDO_Cold = NumOfDO_Cold + (int)numericStepCold.Value;
                    relay2pk = relay2pk + (int)numericStepCold.Value;
                    SocketRelay2pk = SocketRelay2pk + (int)numericStepCold.Value;
                }

                // ���������
                if (checkBoxDraining.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 2;
                    NumOfDO_Dry = NumOfDO_Dry + 1;
                    relay2pk = relay2pk + 1;
                    SocketRelay2pk = SocketRelay2pk + 1;
                }

                // �����������
                if (checkBoxHumOnOff.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 2;
                    NumOfDO_Hum = NumOfDO_Hum + 1;
                    relay2pk = relay2pk + 1;
                    SocketRelay2pk = SocketRelay2pk + 1;
                }

                // ��������� �������� �������
                if ((int.Parse(NumOfAirValve.Text)) > 0)
                {
                    blackTerminal2_5 = blackTerminal2_5 + ((int.Parse(NumOfAirValve.Text) * 2));
                    blueTerminal2_5 = blueTerminal2_5 + ((int.Parse(NumOfAirValve.Text) * 2));
                    NumOfDO_AirValve = NumOfDO_AirValve + int.Parse(NumOfAirValve.Text);
                    relay2pk = relay2pk + (int.Parse(NumOfAirValve.Text));
                    SocketRelay2pk = SocketRelay2pk + (int.Parse(NumOfAirValve.Text));
                }

           
                // ��������� �����������
                if (checkBoxSupply.Checked)
                {
                    NumOfDO_MotorSup = 1;
                    if ((Reserve.Checked) || (duobleReserve.Checked)) 
                    {
                        NumOfDO_MotorSup = 2;
                    }
                }

                // �������� �����������
                if (checkBoxExh.Checked)
                {
                    NumOfDO_MotorExh = 1;
                    if ((Reserve.Checked) || (duobleReserve.Checked))
                    {
                        NumOfDO_MotorExh = 2;
                    }
                }

                NumOfDO_Motor = NumOfDO_MotorSup + NumOfDO_MotorExh;
                NumOfDO = NumOfDO_Motor + NumOfDO_AirValve + NumOfDO_Hum + NumOfDO_Dry + NumOfDO_Cold + NumOfDO_El;

                // ��������� DI

                // ���� �������� ��������
                if ((diffPressEnableSup.Checked) & ((Reserve.Checked) || (duobleReserve.Checked)))
                {
                    blackTerminal2_5 = blackTerminal2_5 + 4;
                    relay2pk = relay2pk + 2;
                    SocketRelay2pk = SocketRelay2pk + 2;
                }
                if ((diffPressEnableSup.Checked) & ((!Reserve.Checked) || (!duobleReserve.Checked)))
                {
                    blackTerminal2_5 = blackTerminal2_5 + 2;
                    relay2pk = relay2pk + 1;
                    SocketRelay2pk = SocketRelay2pk + 1;
                }

                // ����������� � ������� � ����������������
                //    if ((int.Parse(NumOfAirValveExh.Text)) > 0)
                //    {
                //        blackTerminal2_5 = blackTerminal2_5 + ((int.Parse(NumOfAirValveExh.Text) * 2));
                //        blueTerminal2_5 = blueTerminal2_5 + ((int.Parse(NumOfAirValveExh.Text) * 2));
                //        NumOfDO = NumOfDO + 1;
                //    }

                // �������������
                if (WithOutRegulation.Checked)
                {
                    if (AmpsSup < 9)
                    {
                        NameContactor = "��������� ��� 9� 1NO 230� �� EKF AVERES �������: ctr-s-9-10-230-av";

                    }
                    else if (AmpsSup < 12)
                    {
                        NameContactor = "��������� ��� 12� 1NO 230� �� EKF AVERES �������: ctr-s-12-10-230-av";
                    }
                    else if (AmpsSup < 18)
                    {
                        NameContactor = "��������� ��� 18� 1NO 230� �� EKF AVERES �������: ctr-s-18-10-230-av";
                    }
                    else if (AmpsSup < 22)
                    {
                        NameContactor = "��������� ��� 22� 1NO 230� �� EKF AVERES �������: ctr-s-22-10-230-av";
                    }
                    else if (AmpsSup < 25)
                    {
                        NameContactor = "��������� ��� 25� 230� �� EKF AVERES �������: ctr-s-25-00-230-av";
                    }
                    else if (AmpsSup < 30)
                    {
                        NameContactor = "��������� ��� 30� 230� �� EKF AVERES �������: ctr-s-30-00-230-av";
                    }
                    else if (AmpsSup < 32)
                    {
                        NameContactor = "��������� ��� 32� 230� �� EKF AVERES �������: ctr-s-32-00-230-av";
                    }
                    else if (AmpsSup < 38)
                    {
                        NameContactor = "��������� ��� 38� 230� �� EKF AVERES �������: ctr-s-40-00-230-av";
                    }
                    else if (AmpsSup < 50)
                    {
                        NameContactor = "��������� ��� 50� 230� �� EKF AVERES �������: ctr-s-50-00-230-av";
                    }
                    else if (AmpsSup < 60)
                    {
                        NameContactor = "��������� ��� 60� 230� �� EKF AVERES �������: ctr-s-60-00-230-av";
                    }
                    else if (AmpsSup < 65)
                    {
                        NameContactor = "��������� ��� 65� 230� �� EKF AVERES �������: ctr-s-70-00-230-av";
                    }
                    else if (AmpsSup < 80)
                    {
                        NameContactor = "��������� ��� 80� 230� �� EKF AVERES �������: ctr-s-80-00-230-av";
                    }
                    else if (AmpsSup < 90)
                    {
                        NameContactor = "��������� ��� 90� 230� �� EKF AVERES �������: ctr-s-90-00-230-av";
                    }
                    else if (AmpsSup < 100)
                    {
                        NameContactor = "��������� ��� 100� 230� �� EKF AVERES �������: ctr-s-100-00-230-av";
                    }
                    if ((Reserve.Checked) || (duobleReserve.Checked))
                    {
                        contacorSup = contacorSup + 2;
                    }
                    else if (WithOutReserv.Checked)
                    {
                        contacorSup = contacorSup + 1;
                    }

                    NumOfDO = NumOfDO + 1;
                }
                if (WithVFD.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 6;
                    gndTerminal2_5 = gndTerminal2_5 + 1;
                    NumOfDO = NumOfDO + 1;
                    NumOfDI = NumOfDI + 1;
                    NumOfAO = NumOfAO + 1;
                }
                if (SoftStarter.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 4;
                    NumOfDO = NumOfDO + 1;
                }
                if (Transform.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 2;
                    NumOfDO = NumOfDO + 1;
                }
                if (Potenciometr.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 2;
                    NumOfDO = NumOfDO + 1;
                }
                if (ECMotor.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + 6;
                    gndTerminal2_5 = gndTerminal2_5 + 1;
                    NumOfDO = NumOfDO + 1;
                    NumOfDI = NumOfDI + 1;
                    NumOfAO = NumOfAO + 1;
                }

                // ��������� ��

                // ������� �����������
                if (checkBoxHeatAO.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + (((int)numericHeatStep.Value) * 2);
                    gndTerminal2_5 = gndTerminal2_5 + ((int)numericHeatStep.Value);
                    NumOfAO = NumOfAO + ((int)numericHeatStep.Value);
                }

                // ���������� ����������
                if (checkBoxOil.Checked)
                {
                    blackTerminal2_5 = blackTerminal2_5 + (((int)numericOilCold.Value) * 2);
                    gndTerminal2_5 = gndTerminal2_5 + (int)numericOilCold.Value;
                    NumOfAO = NumOfAO + ((int)numericOilCold.Value);
                }

                // �������� �����������
                if (checkBoxGas.Checked)
                {
                    blackTerminal2_5 = (blackTerminal2_5 + 2);
                    gndTerminal2_5 = gndTerminal2_5 + 1;
                    NumOfAO = NumOfAO + 1;
                }

                // ������������
                if (checkBoxRecirc.Checked)
                {
                    blackTerminal2_5 = (blackTerminal2_5 + 2);
                    gndTerminal2_5 = gndTerminal2_5 + 1;
                    NumOfAO = NumOfAO + 1;
                }


                // ����������� �����
                Excel.Range range3 = sheet.get_Range("B2", "J2");
                range3.Merge(Type.Missing);
                Excel.Range range2 = sheet.get_Range("B3", "J3");
                range2.Merge(Type.Missing);
                Excel.Range range1 = sheet.get_Range("B4", "J4");
                range1.Merge(Type.Missing);
                Excel.Range range0 = sheet.get_Range("B1", "J1");
                Excel.Range range7 = sheet.get_Range("A7", "J7");
                range7.Merge(Type.Missing);
                range0.Merge(Type.Missing);
                Excel.Range range8 = sheet.get_Range("A8", "J8");
                range8.Merge(Type.Missing);
                Excel.Range range9 = sheet.get_Range("A9", "J9");
                range9.Merge(Type.Missing);
                Excel.Range range10 = sheet.get_Range("A10", "J10");
                range10.Merge(Type.Missing);
                Excel.Range range11 = sheet.get_Range("A11", "J11");
                range11.Merge(Type.Missing);
                Excel.Range range12 = sheet.get_Range("A12", "J12");
                range12.Merge(Type.Missing);
                Excel.Range range13 = sheet.get_Range("A13", "J13");
                range13.Merge(Type.Missing);
                Excel.Range range14 = sheet.get_Range("A14", "J14");
                range14.Merge(Type.Missing);
                Excel.Range range15 = sheet.get_Range("A15", "J15");
                range15.Merge(Type.Missing);
                Excel.Range range16 = sheet.get_Range("A16", "J16");
                range16.Merge(Type.Missing);
                Excel.Range range17 = sheet.get_Range("A17", "J17");
                range17.Merge(Type.Missing);
                Excel.Range range18 = sheet.get_Range("A18", "J18");
                range18.Merge(Type.Missing);
                Excel.Range range19 = sheet.get_Range("A19", "J19");
                range19.Merge(Type.Missing);
                Excel.Range range20 = sheet.get_Range("A20", "J20");
                range20.Merge(Type.Missing);
                Excel.Range range21 = sheet.get_Range("A21", "J21");
                range21.Merge(Type.Missing);
                Excel.Range range22 = sheet.get_Range("A22", "J22");
                range22.Merge(Type.Missing);
                Excel.Range range23 = sheet.get_Range("A23", "J23");
                range23.Merge(Type.Missing);
                Excel.Range range24 = sheet.get_Range("A24", "J24");
                range24.Merge(Type.Missing);
                Excel.Range range25 = sheet.get_Range("A25", "J25");
                range25.Merge(Type.Missing);
                Excel.Range range26 = sheet.get_Range("A26", "J26");
                range26.Merge(Type.Missing);
                Excel.Range range27 = sheet.get_Range("A27", "J27");
                range27.Merge(Type.Missing);
                Excel.Range range28 = sheet.get_Range("A28", "J28");
                range28.Merge(Type.Missing);
                Excel.Range range29 = sheet.get_Range("A29", "J29");
                range29.Merge(Type.Missing);
                Excel.Range range30 = sheet.get_Range("A30", "J30");
                range30.Merge(Type.Missing);

                //�������� ����� (������� �����)
                sheet.Name = "������������";
                sheet.Range["A1"].Value = "��������";
                sheet.Range["B1"].Value = textBox1.Text;
                sheet.Range["A2"].Value = "������";
                sheet.Range["B2"].Value = textBox2.Text;
                sheet.Range["A3"].Value = "������������� �� ���������";
                sheet.Range["B3"].Value = textBox3.Text;
                sheet.Range["A4"].Value = "�����������";
                sheet.Range["B4"].Value = textBox4.Text;


                sheet.Range["A7"].Value = "������������";
                sheet.Range["K7"].Value = "����������. ��.";
                sheet.Range["L7"].Value = "����, ���";
                sheet.Range["A8"].Value = cabinet;
                sheet.Range["K8"].Value = piceOfCabinet;
                sheet.Range["A9"].Value = mainSwitch;
                sheet.Range["K9"].Value = piceOfMainSwitch;
              //  sheet.Range["A10"].Value = motorProtector;                         
              //  sheet.Range["K10"].Value = pieceOfMotorProtector + " ��.";           
                sheet.Range["A11"].Value = motorPtotectorSup;
                sheet.Range["K11"].Value = pieceOfMotorPtotectorSup + " ��. ";

                // sheet.Range["A12"].Value = motorPtotectorExh;
                // sheet.Range["K12"].Value = pieceOfMotorPtotectorExh + " ��.";


                sheet.Range["A13"].Value = AutomaticSwitchOnePhase;
                sheet.Range["K13"].Value = piceOfAutomaticSwitchOnePhase;
                sheet.Range["A14"].Value = AutomaticSwitchThreePhase;
                sheet.Range["K14"].Value = piceOfAutomaticSwitchThreePhase;
                // ��� ���������� ������������� ����������
                if (WithOutRegulation.Checked)
                {
                    sheet.Range["A15"].Value = NameContactor;
                    sheet.Range["K15"].Value = contacorSup;
                }
                if (checkBoxEl.Checked)
                {
                    sheet.Range["A16"].Value = NameContactorEl;
                    sheet.Range["K16"].Value = contactorEl;
                }

                sheet.Range["A17"].Value = NameRelay2pk;
                sheet.Range["K17"].Value = relay2pk;
                sheet.Range["A18"].Value = NameSocketRelay2pk;
                sheet.Range["K18"].Value = SocketRelay2pk;

                // ����� ������/������� ���

                sheet.Range["A31"].Value = "DO";
                sheet.Range["A32"].Value = "DI";
                sheet.Range["A33"].Value = "AO";
                sheet.Range["A34"].Value = "AI";

                sheet.Range["K31"].Value = NumOfDO;
                sheet.Range["K32"].Value = NumOfDI;
                sheet.Range["K33"].Value = NumOfAO;
                sheet.Range["K34"].Value = NumOfAI;
             

                sheet.Range["A20"].Value = NameBlackTerminal2_5;
                sheet.Range["K20"].Value = ElblackTerminal2_5 + MotorblackTerminal2_5 + blackTerminal2_5;
                sheet.Range["A21"].Value = NameBlueTerminal2_5;
                sheet.Range["K21"].Value = ElblueTerminal2_5 + MotorblueTerminal2_5;
                sheet.Range["A22"].Value = NameGndTerminal2_5;
                sheet.Range["K22"].Value = ElgndTerminal2_5 + MotorgndTerminal2_5;
                sheet.Range["A23"].Value = NameBlackTerminal4;
                sheet.Range["K23"].Value = ElblackTerminal4 + MotorblackTerminal4;
                sheet.Range["A24"].Value = NameGndTerminal4;
                sheet.Range["K24"].Value = ElgndTerminal4 + MotorgndTerminal4;
                sheet.Range["A25"].Value = NameBlackTerminal6;
                sheet.Range["K25"].Value = ElblackTerminal6 + MotorblackTerminal6;
                sheet.Range["A26"].Value = NameGndTerminal6;
                sheet.Range["K26"].Value = ElgndTerminal6 + MotorgndTerminal6;
                sheet.Range["A27"].Value = NameBlackTerminal10;
                sheet.Range["K27"].Value = ElblackTerminal10 + MotorblackTerminal10;
                sheet.Range["A28"].Value = NameGndTerminal10;
                sheet.Range["K28"].Value = ElgndTerminal10 + MotorgndTerminal10;
                sheet.Range["A29"].Value = NameBlackTerminal16;
                sheet.Range["K29"].Value = ElblackTerminal16 + MotorblackTerminal16;
                sheet.Range["A30"].Value = NameGndTerminal16;
                sheet.Range["K30"].Value = ElgndTerminal16 + MotorgndTerminal16;




                // ��������� �������
                if (checkBoxExh.Checked & !checkBoxSupply.Checked)     /// ���� ��� ������ �� ���, ������� ����� ��� ����� else if
                {

                    /* ����� ���������� ��� ������ ����������� ���� */
                    if (ThreePhaseExh.Checked)
                    {
                        double pwr = Double.Parse(PoweExhMain.Text);
                        if ((ReserveExh.Checked) ^ (duobleReserveExh.Checked))
                        {
                            ComAmps = (pwr / 380) * 2;
                            AmpsExh = pwr / 380;
                            //  test.Text = Amps.ToString();
                            //  test1.Text = ComAmps.ToString();
                        }
                        else
                        {
                            ComAmps = pwr / 380;
                            AmpsExh = ComAmps;
                            // test.Text = ComAmps.ToString();
                            //  test1.Text = ComAmps.ToString();
                        }
                    }
                    if (OneNumPhase.Checked)
                    {
                        double pwr = Double.Parse(PoweExhMain.Text);
                        if ((ReserveExh.Checked) ^ (duobleReserveExh.Checked))
                        {
                            ComAmps = (pwr / 220) * 2;
                            AmpsExh = pwr / 220;
                            //test.Text = ComAmps.ToString();
                            // test1.Text = ComAmps.ToString();
                        }
                        else
                            ComAmps = pwr / 220;
                        AmpsExh = ComAmps;
                        //  test.Text = Amps.ToString();
                        //  test1.Text = ComAmps.ToString();
                    }


                    { /* ����� ����� */
                        if (ComAmps < 83)
                        { cabinet = "������ ������� ST ��� �/� 800x600x250 �������: R5ST0869WMP "; }
                        else if (ComAmps < 200)
                        { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1000 x 600 x 250 �� (� � � � �) �������: R5ST1069"; }
                        else if (ComAmps < 400)
                        { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1200 x 800 x 300 �� (� � � � �) �������: R5ST1283"; }
                        else if (ComAmps < 630)
                        { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1400 x 800 x 300 �� (� � � � �) �������: R5ST1483"; }
                        else if (ComAmps < 1000)
                        { cabinet = "DKC ��� ��������� ��� 1600�800�400 IP31 ���� ��� ��������� ������ ������������� ���-16.8.4-0 �������: YKM40-1684-31"; }
                    }


                    /* ����� ���������� */
                    {
                        if (ComAmps < 40)
                        { mainSwitch = "��������� 40A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 40 - 3p - f"; }
                        else if (ComAmps < 63)
                        { mainSwitch = "��������� 63A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 63 - 3p - f"; }
                        else if (ComAmps < 83)
                        { mainSwitch = "��������� 80A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF PROxima �������: tb - 80 - 3p - f"; }
                        else if (ComAmps < 160)
                        { mainSwitch = "��������� 160A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 160 - 3p"; }
                        else if (ComAmps < 200)
                        { mainSwitch = "��������� 200A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 200 - 3p"; }
                        else if (ComAmps < 250)
                        { mainSwitch = "��������� 250A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 250 - 3p"; }
                        else if (ComAmps < 315)
                        { mainSwitch = "��������� 315A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 315 - 3p"; }
                        else if (ComAmps < 400)
                        { mainSwitch = "��������� 400A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 400 - 3p"; }
                        else if (ComAmps < 630)
                        { mainSwitch = "��������� 630A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 630 - 3p"; }
                        else if (ComAmps < 800)
                        { mainSwitch = "��������� 800A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 800 - 3p"; }
                        else if (ComAmps < 1000)
                        { mainSwitch = "��������� 1000A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 1000 - 3p"; }
                        if (ComAmps > 200)
                        {
                            piceOfHandleSwitch = piceOfMainSwitch;
                        }
                    }

                    /* ����� �������� ������ ���������*/
                    {
                        if (AmpsExh < 0.63)
                        {
                            motorProtector = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                        }
                        else if (AmpsExh < 1.0)
                        {
                            motorProtector = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                        }
                        else if (AmpsExh < 1.2)
                        {
                            motorProtector = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                        }
                        else if (AmpsExh < 2.2)
                        {
                            motorProtector = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                        }
                        else if (AmpsExh < 3.6)
                        {
                            motorProtector = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                        }
                        else if (AmpsExh < 5.6)
                        {
                            motorProtector = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                        }
                        else if (AmpsExh < 9)
                        {
                            motorProtector = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                        }
                        else if (AmpsExh < 13.0)
                        {
                            motorProtector = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                        }
                        else if (AmpsExh < 17.0)
                        {
                            motorProtector = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsExh < 22)
                        {
                            motorProtector = "������� ����� ��������� GV2P 17-23 � EKF PROxima �������: gv2p21 - pro";
                        }
                        else if (AmpsExh < 24.0)
                        {
                            motorProtector = "������� ����� ��������� GV2P 20-25 � EKF PROxima �������: gv2p22 - pro";
                        }
                        else if (AmpsExh < 31.0)
                        {
                            motorProtector = "������� ����� ��������� GV2P 24-32 � EKF PROxima �������: gv2p32 - pro";
                        }

                        // ���������� ��������� ������ ����������

                        if ((ReserveExh.Checked) || (duobleReserveExh.Checked))
                        {
                            pieceOfMotorProtector = 2;
                        }
                        else if (WithOutReservExh.Checked)
                            pieceOfMotorProtector = 1;
                    }

                    /*����� ����������� ��������������� ����������� */
                    {
                        AutomaticSwitchOnePhase = "����������� �������������� AV-6 1P 10A (C) 6kA EKF AVERES �������: mcb6 - 1 - 10C - av";
                        piceOfAutomaticSwitchOnePhase = piceOfAutomaticSwitchOnePhase + 1;
                        /* ����� ����������� ��������������� ����������� */
                    }


                    /* EXCEL � �������� �������� ����� */

                    /* ������ � ������ */
                    // ��������� ����� ������
                    //��������� ����������
                    //        Excel.Application app = new Excel.Application
                    //       {
                    //            //���������� Excel
                    //            Visible = true,
                    //            //���������� ������ � ������� �����
                    //            SheetsInNewWorkbook = 2
                    //        };
                    //�������� ������� �����
                    //        Excel.Workbook workBook = app.Workbooks.Add(Type.Missing);
                    //��������� ����������� ���� � �����������
                    //        app.DisplayAlerts = false;
                    //�������� ������ ���� ��������� (���� ���������� � 1)
                    //        Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);



                    //�������� ����� (������� �����)
                    sheet.Name = "������������";
                    sheet.Range["A1"].Value = "��������";
                    sheet.Range["B1"].Value = textBox1.Text;
                    sheet.Range["A2"].Value = "������";
                    sheet.Range["B2"].Value = textBox2.Text;
                    sheet.Range["A3"].Value = "������������� �� ���������";
                    sheet.Range["B3"].Value = textBox3.Text;
                    sheet.Range["A4"].Value = "�����������";
                    sheet.Range["B4"].Value = textBox4.Text;


                    sheet.Range["A7"].Value = "������������";
                    sheet.Range["K7"].Value = "����������";
                    sheet.Range["L7"].Value = "����, ���";
                    sheet.Range["A8"].Value = cabinet;
                    sheet.Range["K8"].Value = piceOfCabinet + " ��.";
                    sheet.Range["A9"].Value = mainSwitch;
                    sheet.Range["K9"].Value = piceOfMainSwitch + " ��.";
                    //   sheet.Range["A10"].Value = motorProtector;
                    //   sheet.Range["K10"].Value = pieceOfMotorProtector + " ��.";

                    // sheet.Range["A11"].Value = motorPtotectorSup;
                    // sheet.Range["K11"].Value = pieceOfMotorPtotectorSup + " ��.";

                    sheet.Range["A12"].Value = motorPtotectorExh;
                    sheet.Range["K12"].Value = pieceOfMotorPtotectorExh + " ��.";


                    sheet.Range["A13"].Value = AutomaticSwitchOnePhase;
                    sheet.Range["K13"].Value = piceOfAutomaticSwitchOnePhase + " ��.";
                    sheet.Range["A14"].Value = AutomaticSwitchThreePhase;
                    sheet.Range["K14"].Value = piceOfAutomaticSwitchThreePhase + " ��.";



                }
                // ��������� �������� �������� �������         
                if (checkBoxSupply.Checked & checkBoxExh.Checked)

                {
                    /* ����� ���������� ��� ������ ����������� ���� */
                    if (checkBoxSupply.Checked & checkBoxExh.Checked)
                    {
                        double pwrExh = Double.Parse(PoweExhMain.Text);
                        double pwrSup = Double.Parse(PoweSupMain.Text);
                        ComAmps = (pwrExh / 380) + (pwrSup / 380);
                        AmpsExh = (pwrExh / 380);
                        AmpsSup = (pwrSup / 380);

                        if ((ReserveExh.Checked || duobleReserveExh.Checked) & (Reserve.Checked || duobleReserve.Checked))
                        {
                            ComAmps = ((pwrExh / 380) + (pwrSup / 380)) * 2;
                        }


                    }
                    if (OneNumPhase.Checked)
                    {
                        double pwrExh = Double.Parse(PoweExhMain.Text);
                        double pwrSup = Double.Parse(PoweSupMain.Text);

                        if ((ReserveExh.Checked & duobleReserveExh.Checked))
                        {
                            ComAmps = ((pwrExh / 220) + (pwrSup / 220)) * 2;
                            Amps = (pwrExh / 220);
                            //test.Text = ComAmps.ToString();
                            // test1.Text = ComAmps.ToString();
                        }
                        else
                            ComAmps = (pwrExh / 220) + (pwrSup / 220);
                        Amps = pwrExh / 220;
                        //  test.Text = Amps.ToString();
                        //  test1.Text = ComAmps.ToString();
                    }


                    /* ����� ����� */
                    if (ComAmps < 83)
                    { cabinet = "������ ������� ST ��� �/� 800x600x250 �������: R5ST0869WMP "; }
                    else if (ComAmps < 200)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1000 x 600 x 250 �� (� � � � �) �������: R5ST1069"; }
                    else if (ComAmps < 400)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1200 x 800 x 300 �� (� � � � �) �������: R5ST1283"; }
                    else if (ComAmps < 630)
                    { cabinet = "DKC ������ ������� �������� ����� ST � �/� ������: 1400 x 800 x 300 �� (� � � � �) �������: R5ST1483"; }
                    else if (ComAmps < 1000)
                    { cabinet = "DKC ��� ��������� ��� 1600�800�400 IP31 ���� ��� ��������� ������ ������������� ���-16.8.4-0 �������: YKM40-1684-31"; }



                    /* ����� ���������� */

                    if (ComAmps < 40)
                    { mainSwitch = "��������� 40A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 40 - 3p - f"; }
                    else if (ComAmps < 63)
                    { mainSwitch = "��������� 63A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF �������: tb - 63 - 3p - f"; }
                    else if (ComAmps < 83)
                    { mainSwitch = "��������� 80A 3P c ��������� ���������� ��� ������ ��������� TwinBlock EKF PROxima �������: tb - 80 - 3p - f"; }
                    else if (ComAmps < 160)
                    { mainSwitch = "��������� 160A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 160 - 3p"; }
                    else if (ComAmps < 200)
                    { mainSwitch = "��������� 200A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 200 - 3p"; }
                    else if (ComAmps < 250)
                    { mainSwitch = "��������� 250A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 250 - 3p"; }
                    else if (ComAmps < 315)
                    { mainSwitch = "��������� 315A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 315 - 3p"; }
                    else if (ComAmps < 400)
                    { mainSwitch = "��������� 400A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 400 - 3p"; }
                    else if (ComAmps < 630)
                    { mainSwitch = "��������� 630A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 630 - 3p"; }
                    else if (ComAmps < 800)
                    { mainSwitch = "��������� 800A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 800 - 3p"; }
                    else if (ComAmps < 1000)
                    { mainSwitch = "��������� 1000A 3P ��� �������� ���������� TwinBlock EKF PROxima �������: tb - s - 1000 - 3p"; }
                    if (ComAmps > 200)
                    {
                        piceOfHandleSwitch = piceOfMainSwitch;
                    }


                    /* ����� �������� ������ ���������*/
                    {
                        // ������� ���������� �����������

                        if (AmpsSup < 0.63) // & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                        }
                        else if (AmpsSup < 1.0) // & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                        }
                        else if (AmpsSup < 1.2)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                            //   motorPtotectorExh  = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                        }
                        else if (AmpsSup < 2.2)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                        }
                        else if (AmpsSup < 3.6)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                        }
                        else if (AmpsSup < 5.6)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                        }
                        else if (AmpsSup < 9)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                            //  motorPtotectorExh = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                        }
                        else if (AmpsSup < 13.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                        }
                        else if (AmpsSup < 17.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                            //    motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsSup < 22)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 17-23 � EKF PROxima �������: gv2p21 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsSup < 24.0) //&(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 20-25 � EKF PROxima �������: gv2p22 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsSup < 31.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            motorPtotectorSup = "������� ����� ��������� GV2P 24-32 � EKF PROxima �������: gv2p32 - pro";
                            //   motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }

                        // ������� ��������� �����������

                        if (AmpsExh < 0.63) // & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            //  motorPtotectorSup = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 0,4-0,63 � EKF PROxima �������: gv2p04 - pro";
                        }
                        else if (AmpsExh < 1.0) // & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 0,63-1,0 � EKF PROxima �������: gv2p05 - pro";
                        }
                        else if (AmpsExh < 1.2)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            //  motorPtotectorSup = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 1,0-1,6 � EKF PROxima �������: gv2p06 - pro";
                        }
                        else if (AmpsExh < 2.2)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            //  motorPtotectorSup = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 1,6-2,5 � EKF PROxima �������: gv2p07 - pro";
                        }
                        else if (AmpsExh < 3.6)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 2,5-4 � EKF PROxima �������: gv2p08 - pro";
                        }
                        else if (AmpsExh < 5.6)// & (checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 4-6,3 � EKF PROxima �������: gv2p10 - pro";
                        }
                        else if (AmpsExh < 9)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 6-10 � EKF PROxima �������: gv2p14 - pro";
                        }
                        else if (AmpsExh < 13.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            //motorPtotectorSup = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 9-14 � EKF PROxima �������: gv2p16 - pro";
                        }
                        else if (AmpsExh < 17.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsExh < 22)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 17-23 � EKF PROxima �������: gv2p21 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (AmpsExh < 24.0) //&(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 20-25 � EKF PROxima �������: gv2p22 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";
                        }
                        else if (Amps < 31.0)// &(checkBoxSupply.Checked & checkBoxExh.Checked))
                        {
                            // motorPtotectorSup = "������� ����� ��������� GV2P 24-32 � EKF PROxima �������: gv2p32 - pro";
                            motorPtotectorExh = "������� ����� ��������� GV2P 13-18 � EKF PROxima �������: gv2p20 - pro";


                        }

                        // ���������� ��������� ������ ����������

                        if (((ReserveExh.Checked) || (duobleReserveExh.Checked)) & ((Reserve.Checked) || (duobleReserve.Checked)))
                        {
                            pieceOfMotorPtotectorExh = 2;
                            pieceOfMotorPtotectorSup = 2;
                        }
                        else if (((!ReserveExh.Checked) || (!duobleReserveExh.Checked)) & ((Reserve.Checked) || (duobleReserve.Checked)))
                        {
                            pieceOfMotorPtotectorExh = 1;
                            pieceOfMotorPtotectorSup = 2;
                        }
                        else if (((ReserveExh.Checked) || (duobleReserveExh.Checked)) & (!(Reserve.Checked) || (!duobleReserve.Checked)))
                        {
                            pieceOfMotorPtotectorExh = 2;
                            pieceOfMotorPtotectorSup = 1;
                        }

                        else if (((WithOutReserv.Checked) & (WithOutReservExh.Checked)) || ((!ReserveExh.Checked) || (!duobleReserveExh.Checked)) & ((!Reserve.Checked) || (!duobleReserve.Checked)))
                        {
                            pieceOfMotorPtotectorExh = 1;
                            pieceOfMotorPtotectorSup = 1;
                        }

                    }

                    /*����� ����������� ��������������� ����������� */
                    {
                        AutomaticSwitchOnePhase = "����������� �������������� AV-6 1P 10A (C) 6kA EKF AVERES �������: mcb6 - 1 - 10C - av";
                        piceOfAutomaticSwitchOnePhase = piceOfAutomaticSwitchOnePhase + 1;
                        /* ����� ����������� ��������������� ����������� */
                    }



                    /* EXCEL � �������� �������� ����� */

                    /* ������ � ������ */
                    // ��������� ����� ������
                    //��������� ����������
                    //                Excel.Application app = new Excel.Application
                    //                {
                    //                    //���������� Excel
                    //                    Visible = true,
                    //                    //���������� ������ � ������� �����
                    //                   SheetsInNewWorkbook = 2
                    //               };
                    //                //�������� ������� �����
                    //               Excel.Workbook workBook = app.Workbooks.Add(Type.Missing);
                    //               //��������� ����������� ���� � �����������
                    //               app.DisplayAlerts = false;
                    //               //�������� ������ ���� ��������� (���� ���������� � 1)
                    //               Excel.Worksheet sheet = (Excel.Worksheet)app.Worksheets.get_Item(1);


                    //�������� ����� (������� �����)
                    sheet.Name = "������������";
                    sheet.Range["A1"].Value = "��������";
                    sheet.Range["B1"].Value = textBox1.Text;
                    sheet.Range["A2"].Value = "������";
                    sheet.Range["B2"].Value = textBox2.Text;
                    sheet.Range["A3"].Value = "������������� �� ���������";
                    sheet.Range["B3"].Value = textBox3.Text;
                    sheet.Range["A4"].Value = "�����������";
                    sheet.Range["B4"].Value = textBox4.Text;


                    sheet.Range["A7"].Value = "������������";
                    sheet.Range["K7"].Value = "����������";
                    sheet.Range["L7"].Value = "����, ���";
                    sheet.Range["A8"].Value = cabinet;
                    sheet.Range["K8"].Value = piceOfCabinet + " ��.";
                    sheet.Range["A9"].Value = mainSwitch;
                    sheet.Range["K9"].Value = piceOfMainSwitch + " ��.";
                    //   sheet.Range["A10"].Value = motorProtector;
                    //   sheet.Range["K10"].Value = pieceOfMotorProtector + " ��.";

                    sheet.Range["A11"].Value = motorPtotectorSup;
                    sheet.Range["K11"].Value = pieceOfMotorPtotectorSup + " ��.";

                    sheet.Range["A12"].Value = motorPtotectorExh;
                    sheet.Range["K12"].Value = pieceOfMotorPtotectorExh + " ��.";


                    sheet.Range["A13"].Value = AutomaticSwitchOnePhase;
                    sheet.Range["K13"].Value = piceOfAutomaticSwitchOnePhase + " ��.";
                    sheet.Range["A14"].Value = AutomaticSwitchThreePhase;
                    sheet.Range["K14"].Value = piceOfAutomaticSwitchThreePhase + " ��.";


                }



            }

            // ����� �����

            ComblackTerminal2_5 = ElblackTerminal2_5 + MotorblackTerminal2_5;
            ComblueTerminal2_5 = ElblueTerminal2_5 + MotorblueTerminal2_5;
            ComgndTerminal2_5 = ElgndTerminal2_5 + MotorgndTerminal2_5;
            ComblackTerminal4 = ElblackTerminal4 + MotorblackTerminal4;
            ComgndTerminal4 = ElgndTerminal4 + MotorgndTerminal4;
            ComblackTerminal6 = ElblackTerminal6 + MotorblackTerminal6;
            ComgndTerminal6 = ElgndTerminal6 + MotorgndTerminal6;
            ComblackTerminal10 = ElblackTerminal10 + MotorblackTerminal10;
            ComgndTerminal10 = ElgndTerminal10 + MotorgndTerminal10;
            ComblackTerminal16 = ElblackTerminal16 + MotorblackTerminal16;
            ComgndTerminal16 = ElgndTerminal16 + MotorgndTerminal16;






        }

        private void numericHeatStep_ValueChanged(object sender, EventArgs e)
        {

        }

        private void checkBox17_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
