using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO.Ports;
using System.IO;
using System.Text.RegularExpressions;

using System.Timers;

using One_Sgp4;
using Microsoft.VisualBasic.FileIO;

namespace OpenBirdsOperation
{
    public partial class operationMain : Form
    {

        public string CommandString
        {
            set
            {
                Command_HEX.Text = value;
            }
            get
            {
                return Command_HEX.Text;
            }
        }

        public string CommentsString
        {
            set
            {
                Comments.Text = value;
            }
            get
            {
                return Comments.Text;
            }
        }

        private SerialPort TncPort = new SerialPort();
        private SerialPort RadioPort = new SerialPort();
        private SerialPort RotatorPort = new SerialPort();

        private bool TncPortConnected = false;
        private bool RadioPortConnected = false;
        private bool RotatorPortConnected = false;
        private bool ReceivedDataKissFrameFlag = false;

        int receivedPacketNumber = 0;

        public commandList commandList;

        public string rootPath;
        public DateTime dtAppStart = DateTime.UtcNow;
        public string savingRootPath;
        public string tlePath;
        public string freqPath;

        public double latitude, longitude, height;

        public double ulFreq, dlFreq;
        public double cwFreq;

        public int civAddress;

        public double range;
        public double elevation, azimuth;

        List<Tle> tleList;

        Tle selectedTle;

        System.Windows.Forms.Timer trackingTimer;

        bool uplinkFlag = false;

        int freqAddjustValue = 0;

        double azHomeValue;
        double elHomeValue;

        string satFolderName = "NULL";

        private void prosistelRotatorCommand(double azimuth, double elevation)
        {

            if (!RotatorPort.IsOpen) {
                return;
            }

            if (elevation < 0) {
                elevation = 0;
            }

            byte[] rotatorCMD = new byte[16];

            //azimuth
            rotatorCMD[0]  = 0x02;
            rotatorCMD[1]  = Convert.ToByte('A');    //azimuth
            rotatorCMD[2]  = Convert.ToByte('G');    //GOTO command
            rotatorCMD[3]  = Convert.ToByte((int)(azimuth / 100) % 10 + '0');
            rotatorCMD[4]  = Convert.ToByte((int)(azimuth /  10) % 10 + '0');
            rotatorCMD[5]  = Convert.ToByte((int)(azimuth /   1) % 10 + '0');
            rotatorCMD[6]  = Convert.ToByte((int)(azimuth *  10) % 10 + '0');
            rotatorCMD[7]  = 0x0D;

            //elevation
            rotatorCMD[ 8] = 0x02;
            rotatorCMD[ 9] = Convert.ToByte('B');    //azimuth
            rotatorCMD[10] = Convert.ToByte('G');    //GOTO command
            rotatorCMD[11] = Convert.ToByte((int)(elevation / 100) % 10 + '0');
            rotatorCMD[12] = Convert.ToByte((int)(elevation /  10) % 10 + '0');
            rotatorCMD[13] = Convert.ToByte((int)(elevation /   1) % 10 + '0');
            rotatorCMD[14] = Convert.ToByte((int)(elevation *  10) % 10 + '0');
            rotatorCMD[15] = 0x0D;

//            Console.WriteLine(System.Text.Encoding.Default.GetString(rotatorCMD));

            RotatorPort.Write(rotatorCMD, 0, rotatorCMD.Length);

        }

        string rotatorCMDhist = "";
        private void yaesuRotatorCommand(double azimuth, double elevation)
        {
            if (!RotatorPort.IsOpen)
            {
                return;
            }
            
            if (elevation < 0)
            {
                elevation = 0;
            }
            byte[] rotatorCMD = new byte[9];

            rotatorCMD[0] = Convert.ToByte('W');    //set both angle
            rotatorCMD[1] = Convert.ToByte((int)(azimuth   / 100) % 10 + '0');
            rotatorCMD[2] = Convert.ToByte((int)(azimuth   /  10) % 10 + '0');
            rotatorCMD[3] = Convert.ToByte((int)(azimuth   /   1) % 10 + '0');
            rotatorCMD[4] = Convert.ToByte(' ');
            rotatorCMD[5] = Convert.ToByte((int)(elevation / 100) % 10 + '0');
            rotatorCMD[6] = Convert.ToByte((int)(elevation /  10) % 10 + '0');
            rotatorCMD[7] = Convert.ToByte((int)(elevation /   1) % 10 + '0');
            rotatorCMD[8] = 0x0D;

            if (Encoding.GetEncoding("ascii").GetString(rotatorCMD) != rotatorCMDhist) {     //if the CMD is not same to before sending
                RotatorPort.Write(rotatorCMD, 0, rotatorCMD.Length);
                rotatorCMDhist = Encoding.GetEncoding("ascii").GetString(rotatorCMD);
                Console.WriteLine(rotatorCMDhist);
            }
        }

        private void setSatelliteComboList()
        {
            //Parse tle from file
            if (!System.IO.File.Exists(tlePath))
            {
                MessageBox.Show("TLE file is no exists");
                return;
            }
            tleList = ParserTLE.ParseFile(tlePath);
            SatelliteNameComboBox.Items.Clear();
            SatelliteNameComboBoxSub.Items.Clear();

            foreach (var tle in tleList)
            {
                if (!System.IO.File.Exists(freqPath))
                {
                    SatelliteNameComboBox.Items.Add(tle.getNoradID());
                    SatelliteNameComboBoxSub.Items.Add(tle.getNoradID());
                }
                else {

                    using (var parser = new TextFieldParser(freqPath)) {
                        parser.Delimiters = new string[] { "," };
                        while (!parser.EndOfData)
                        {
                            var fields = parser.ReadFields();
                            if (fields[1] == tle.getNoradID()) {
                                SatelliteNameComboBox.Items.Add(fields[0]);
                                SatelliteNameComboBoxSub.Items.Add(fields[0]);
                                break;
                            }
                        }
                    }
                }

            }
        }

        private void updateSavingPath()
        {
            rootPath = Properties.Settings.Default.savingPath;

            savingRootPath = System.IO.Path.Combine(rootPath, dtAppStart.ToString("yyyyMMdd"));

            if (!Directory.Exists(rootPath))
            {
                Directory.CreateDirectory(rootPath);
            }
            if (!Directory.Exists(savingRootPath))
            {
                Directory.CreateDirectory(savingRootPath);
            }

        }

        public operationMain()
        {
            InitializeComponent();
            TncPort.DataReceived += new SerialDataReceivedEventHandler(DataReceived); //DataReceived is the event handler
            commandList = new commandList();
            commandList.formMain = this;
            commandList.Show();

            updateSavingPath();

        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            //When program is loaded all components related COM port are disenable.
            TncPortComboBox.Enabled = false;
            RadioPortComboBox.Enabled = false;
            TncConnectButton.Enabled = false;
            RadioConnectButton.Enabled = false;
//            TransmitButton.Enabled = false;       test
            CwRadioButton.Enabled = false;
            FmdRadioButton.Enabled = false;

            TncBaudComboBox.Enabled = false;
            RadioBaudComboBox.Enabled = false;

            RotatorPortComboBox.Enabled = false;
            RotatorConnectButton.Enabled = false;
            RotatorBaudComboBox.Enabled = false;

            /*
            SatelliteNameComboBox.Enabled = false;
            buttonMinus500Hz.Enabled = false;
            buttonMinus1kHz.Enabled = false;
            buttonMinus5kHz.Enabled = false;
            buttonPulus500Hz.Enabled = false;
            buttonPulus1kHz.Enabled = false;
            buttonPulus5kHz.Enabled = false;
            */
            savingFolderPathTextBox.Text = savingRootPath;
//            TransmitButton.Enabled = false;
            SaveReceiveDataButton.Enabled = false;


            //Add port list into each COM port ComboBox
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                TncPortComboBox.Items.Add(port);
                RadioPortComboBox.Items.Add(port);
                RotatorPortComboBox.Items.Add(port);
            }

            TncPortComboBox.Enabled = true;
            TncConnectButton.Enabled = true;
            TncBaudComboBox.Enabled = true;
            RadioPortComboBox.Enabled = true;
            RadioConnectButton.Enabled = true;
            RadioBaudComboBox.Enabled = true;
            RotatorPortComboBox.Enabled = true;
            RotatorConnectButton.Enabled = true;
            RotatorBaudComboBox.Enabled = true;

            //Number of COM port is more than 1 -> TNC COM port is enabled
            if (TncPortComboBox.Items.Count >= 1)
            {
                TncPortComboBox.SelectedIndex = 0;
                //From BIRDS-1, for TNC Port, added Radio Port
            }

            //Number of COM port is more than 2 -> Radio COM port is enabled
            if (RadioPortComboBox.Items.Count >= 2)
            {
                RadioPortComboBox.SelectedIndex = 1;
            }

            //Number of COM port is more than 3 -> Rotator COM port is enabled
            if (RotatorPortComboBox.Items.Count >= 3)
            {
                RotatorPortComboBox.SelectedIndex = 2;
            }

            radioModelComboBox.SelectedIndex = 0;
            //            rotatorModelComboBox.SelectedIndex = 0;
            if (0 <= Properties.Settings.Default.rotatorModel && Properties.Settings.Default.rotatorModel < 2)
            {
                rotatorModelComboBox.SelectedIndex = Properties.Settings.Default.rotatorModel;
            }

            tlePath = Properties.Settings.Default.tlePath;
            tlePathTextBox.Text = tlePath;

            freqPath = Properties.Settings.Default.freqPath;
            freqPathTextBox.Text = freqPath;


            SatelliteNameComboBox.Enabled = true;   //test
            setSatelliteComboList();

            latitudeTextBox.Text = Properties.Settings.Default.latitude.ToString("0.0000");
            longitudeTextBox.Text = Properties.Settings.Default.longitude.ToString("0.0000");
            heightTextBox.Text = Properties.Settings.Default.height.ToString("0.0");

            latitude = Convert.ToDouble(latitudeTextBox.Text);
            longitude = Convert.ToDouble(longitudeTextBox.Text);
            height = Convert.ToDouble(heightTextBox.Text);

            civAddressTextBox.Text = Properties.Settings.Default.CIVaddress.ToString("X00");
            civAddress = Properties.Settings.Default.CIVaddress;



            trackingCheckBox.Checked = true;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();

            trackingTimer = new System.Windows.Forms.Timer();
            trackingTimer.Interval = 100;   //every 0.1sec
            trackingTimer.Tick += satTracking;
            trackingTimer.Enabled = true;

            azHomeValue = Properties.Settings.Default.azHome;
            elHomeValue = Properties.Settings.Default.elHome;

            azHomeTextBox.Text = azHomeValue.ToString(".0");
            elHomeTextBox.Text = elHomeValue.ToString(".0");

        }

        public void Transmit(object sender, EventArgs e) {
            Transmit_async();
        }


        public async Task Transmit_async()
        {
            string Sending_command = Command_HEX.Text;
            string Sending_coments = Comments.Text;


            //2021/06/13 only check 14bytes
            if (Regex.IsMatch(Sending_command, ".. .. .. .. .. .. .. .. .. .. ..                      "))
            {
                Console.WriteLine("14bytes command");
                Sending_command = Sending_command.Substring(0, 32);
            }
            else {
                Console.WriteLine("invalid length");
                MessageBox.Show("Invalid length");
                return;
            }

            if (Sending_command == "00 00 00 00 00 00 00 00 00 00 00") {
                DialogResult result = MessageBox.Show(
                    "Is it OK to send an invalid command?\nThe command is all 00",
                    "Invalid command",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Exclamation,
                    MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Cancel) {
                    return;
                }
            }

            TransmitButton.Enabled = false;
            byte[] packet = Sending_command.ToString().Split(' ').Select(s => Convert.ToByte(s, 16)).ToArray();
            byte[] packet_withB = new byte[packet.Length + 1];
            packet_withB[0] = (byte)'B';     //'B' it is for Addnics receiver
            Buffer.BlockCopy(packet, 0, packet_withB, 1, packet.Length);
            byte[] KISSframe_toTNC = AX25packetToKISSframe(packet_withB);
            //            byte[] KISSframe_toTNC = AX25packetToKISSframe(packet);

            FmdRadioButton.Checked = true;
            await Task.Delay(100);

            ChangeUplinkFrequency();
            await Task.Delay(200);

            if (TncPortConnected) {
                TncPort.Write(KISSframe_toTNC, 0, KISSframe_toTNC.Length);
            }

            await Task.Delay(1500);
            ChangeDownlinkFrequency();
            DateTime dt = DateTime.UtcNow;
            string time = dt.ToString("yyyy/MM/dd HH:mm:ss");

//            RxDataTextBox.Text += "#(" + time + ")" + " CMD: " + Sending_command + "\n";
            this.RxDataTextBox.Focus();
            this.RxDataTextBox.AppendText("#(" + time + ")" + " CMD: " + Sending_command + "\n");

            await Task.Delay(5000);
            TransmitButton.Enabled = true;

        }
        private byte[] AX25packetToKISSframe(byte[] ax25packet)
        {
            byte[] KISSframe_temp = new byte[3 + 2 * ax25packet.Length];
            int index2 = 0; //index for ax25packet_encoded 

            if (TncBaudComboBox.Text == "115200") {
                System.Diagnostics.Trace.WriteLine("ADDNICS TNC");
                //                KISSframe_temp[index2++] = 0x42;  //ADDNICS 600bps header
            }

            KISSframe_temp[index2++] = 0xC0;  //KISS header (index2 = 0)
            KISSframe_temp[index2++] = 0x00;    //port number (index2 = 1)

            for (int index1 = 0; index1 < ax25packet.Length; index1++)
            {
                if (ax25packet[index1] == 0xC0)   //0xC0
                {
                    //Replace with 0xDB 0xDC
                    KISSframe_temp[index2++] = 0xDB;
                    KISSframe_temp[index2++] = 0xDC;
                }
                else if (ax25packet[index1] == 0xDB)   //0xDB
                {
                    //Replace with 0xDB 0xDD
                    KISSframe_temp[index2++] = 0xDB;
                    KISSframe_temp[index2++] = 0xDD;
                }
                else KISSframe_temp[index2++] = ax25packet[index1];

            }

            KISSframe_temp[index2++] = 0xC0;    //KISS footer

            byte[] KISSframe = new byte[index2];
            Buffer.BlockCopy(KISSframe_temp, 0, KISSframe, 0, index2);

            for (int i = 0; i < KISSframe.Length; i++) {
                System.Diagnostics.Trace.Write(string.Format("{0,3:X2}", KISSframe[i]));
            }
            System.Diagnostics.Trace.WriteLine("");

            return KISSframe;
        }

        public void ChangeFrequency(int freq)
        {
            //Set radio for command uplink
            if (RadioPortConnected == true)
            {
                //            byte[] frequency_command = { 0xFE, 0xFE, 0x00, 0x9D, 0x00, 0x00, 0x30, 0x31, 0x35, 0x04, 0xFD };   //435.313MHz
                byte[] freqCMD = new byte[11];
                freqCMD[0] = 0xFE;
                freqCMD[1] = 0xFE;
                freqCMD[2] = 0x00;
                freqCMD[3] = BitConverter.GetBytes(civAddress)[0];
                freqCMD[4] = 0x00; //Change freqency command 0x00
                freqCMD[5] = BitConverter.GetBytes(freq / 1 % 10 + freq / 10 % 10 * 16)[0];    //0x12 -> XXXXXXXX12Hz
                freqCMD[6] = BitConverter.GetBytes(freq / 100 % 10 + freq / 1000 % 10 * 16)[0];    //0x34 -> XXXXXX34XXHz
                freqCMD[7] = BitConverter.GetBytes(freq / 10000 % 10 + freq / 100000 % 10 * 16)[0];    //0x56 -> XXXX56XXXXHz
                freqCMD[8] = BitConverter.GetBytes(freq / 1000000 % 10 + freq / 10000000 % 10 * 16)[0];    //0x78 -> XX78XXXXXXHz
                freqCMD[9] = BitConverter.GetBytes(freq / 100000000 % 10 + freq / 1000000000 % 10 * 16)[0];    //0x90 -> 90XXXXXXXXHz
                freqCMD[10] = 0xFD;

//                Console.WriteLine(BitConverter.ToString(freqCMD));
                RadioPort.Write(freqCMD, 0, freqCMD.Length);
            }


        }


        public delegate void UpdateTrackingValueDelegate(int frequency, double azimuth, double elevation);

        public void UpdateTrakingValue(int frequency, double azimuth, double elevation)
        {
            freqLabel.Text = (frequency/1000000.0).ToString("000.000000");
            elevationLabel.Text = elevation.ToString("00.0");
            azimuthLabel.Text = azimuth.ToString("000.0");

            ChangeFrequency(frequency);


            if (elevation > -10)
            {
                if (rotatorModelComboBox.SelectedIndex == 0)
                {
                    prosistelRotatorCommand(azimuth, elevation);
                }
                else {
                    yaesuRotatorCommand(azimuth, elevation);
                }
            }
            else {
                if (rotatorModelComboBox.SelectedIndex == 0)
                {
                    prosistelRotatorCommand(azHomeValue, elHomeValue);
                }
                else
                {
                    yaesuRotatorCommand(azHomeValue, elHomeValue);
                }
            }

        }

        //tracking function that timer call every 1 or 0.1 sec
        public void satTracking(object sender, EventArgs e)
        {

            double deltaT = 0.1;         //deltaT = 0.1s
            double speedOfLight = 299792458;      //unit : m/s
            int correctedFreqency;
            double rangingSpeed;

            UpdateTrackingValueDelegate UTVD = new UpdateTrackingValueDelegate(UpdateTrakingValue);

            if (selectedTle == null) {
                this.Invoke(UTVD, 0, azHomeValue, elHomeValue);
                return;
            }

            if (!trackingCheckBox.Checked) {
                if (uplinkFlag)
                {
                    //FM uplink
                    correctedFreqency = (int)((ulFreq + freqAddjustValue));
                }
                else
                {
                    if (CwRadioButton.Checked)
                    {
                        //CW listning
                        correctedFreqency = (int)((cwFreq + freqAddjustValue));
                    }
                    else
                    {
                        //FM downlink
                        correctedFreqency = (int)((dlFreq + freqAddjustValue));
                    }
                }
                this.Invoke(UTVD, correctedFreqency, azHomeValue, elHomeValue);
                return;
            }

            //Create Time points
            EpochTime currentTime = new EpochTime(DateTime.UtcNow);
            EpochTime nextTime = new EpochTime(DateTime.UtcNow.AddMilliseconds(deltaT * 1000));
            //Calculate Satellite Position and Speed
            One_Sgp4.Sgp4 sgp4Propagator = new Sgp4(selectedTle, Sgp4.wgsConstant.WGS_84);
            //set calculation parameters StartTime, EndTime and caclulation steps in minutes
            sgp4Propagator.runSgp4Cal(currentTime,nextTime,1/60.0 * deltaT);  //100msec
            List<One_Sgp4.Sgp4Data> resultDataList = new List<Sgp4Data>();
            //Return Results containing satellite Position x,y,z (ECI-Coordinates in Km) and Velocity x_d, y_d, z_d (ECI-Coordinates km/s) 
            resultDataList = sgp4Propagator.getResults();

            //Coordinate of an observer on Ground lat, long, height(in meters)
//            One_Sgp4.Coordinate observer = new Coordinate(latitude, longitude, height);
            //Coordinate of an observer on Ground lat, long, height(in meters)

//            One_Sgp4.Coordinate observer = new Coordinate(33.8925, 130.8402, 42.0);
            One_Sgp4.Coordinate observer = new Coordinate(latitude, longitude, height);



            //Calculate Sperical Coordinates from an Observer to Satellite
            //returns 3D-Point with range(km), azimuth(radians), elevation(radians) to the Satellite
            One_Sgp4.Point3d spherical = One_Sgp4.SatFunctions.calcSphericalCoordinate(observer, currentTime, resultDataList[0]);
            One_Sgp4.Point3d nextSpherical = One_Sgp4.SatFunctions.calcSphericalCoordinate(observer, currentTime, resultDataList[1]);

            rangingSpeed = (nextSpherical.x - spherical.x) * 1000 / deltaT;  //unit : m/s

            if (uplinkFlag)
            {
                //FM uplink
                correctedFreqency = (int)((ulFreq - freqAddjustValue) * (1 + rangingSpeed / speedOfLight));

                //by assuming freq addjust is coused by doppler shift and orbit calculating error. freqAddjustValue should be opposit.
            }
            else
            {
                if (CwRadioButton.Checked)
                {
                    //CW listning
                    correctedFreqency = (int)((cwFreq + freqAddjustValue) * (1 - rangingSpeed / speedOfLight));
                }
                else
                {
                    //FM downlink
                    correctedFreqency = (int)((dlFreq + freqAddjustValue) * (1 - rangingSpeed / speedOfLight));
                }
            }



            this.Invoke(UTVD, correctedFreqency, spherical.y,  spherical.z);

        }

        public void ChangeUplinkFrequency()
        {
            uplinkFlag = true;
        }

        public void ChangeDownlinkFrequency()
        {
            uplinkFlag = false;
        }

        public void ChangeCW()
        {
            //Set radio for command uplink
            if (RadioPortConnected == true)
            {
                byte[] radioMode_CW = { 0xFE, 0xFE, 0x00, 0x7C, 0x01, 0x03, 0x01, 0xFD, 0xFE, 0xFE, 0x00, 0x7C, 0x1A, 0x06, 0x00, 0x00, 0xFD };   //CW, Data mode off
                radioMode_CW[ 3] = BitConverter.GetBytes(civAddress)[0];
                RadioPort.Write(radioMode_CW, 0, radioMode_CW.Length);
            }
        }

        public void ChangeFMD()
        {
            //Set radio for command uplink
            if (RadioPortConnected == true)
            {
                byte[] radioMode_FMD = { 0xFE, 0xFE, 0x00, 0x7C, 0x01, 0x05, 0x01, 0xFD, 0xFE, 0xFE, 0x00, 0x7C, 0x1A, 0x06, 0x01, 0x01, 0xFD };  //FM, Data mode on
                radioMode_FMD[3] = BitConverter.GetBytes(civAddress)[0];
                RadioPort.Write(radioMode_FMD, 0, radioMode_FMD.Length);
            }
        }

        private void TncConnectButton_Click(object sender, EventArgs e)
        {

            if (!TncPortConnected)
            {
                TncPort.PortName = TncPortComboBox.Text;
                TncPort.BaudRate = Convert.ToInt32(TncBaudComboBox.SelectedItem);
                TncPort.DataBits = 8;
                TncPort.Parity = Parity.None;
                TncPort.StopBits = StopBits.One;
                TncPort.DtrEnable = true;
                if (Convert.ToInt32(TncBaudComboBox.SelectedItem) != 115200)
                {
                    TncPort.Handshake = Handshake.RequestToSend;
                }
                else {
                    TncPort.Handshake = Handshake.None;
                    System.Diagnostics.Trace.WriteLine("Connect ADDNICS TNC");
                }
                TncPort.Open();
                TncPortConnected = true;
                TransmitButton.Enabled = true;
                TncPortComboBox.Enabled = false;
                TncBaudComboBox.Enabled = false;
                TncConnectButton.Text = "DisConnect";


            }
            else
            {
                TncPort.Close();
                TncPortConnected = false;
                TransmitButton.Enabled = false;
                TncPortComboBox.Enabled = true;
                TncBaudComboBox.Enabled = true;
                TncConnectButton.Text = "Connect";


            }

        }

        private void RadioConnectButton_Click(object sender, EventArgs e)
        {
            if (!RadioPortConnected)
            {
                RadioPort.PortName = RadioPortComboBox.Text;
                RadioPort.BaudRate = Convert.ToInt32(RadioBaudComboBox.SelectedItem);
                RadioPort.DataBits = 8;
                RadioPort.Parity = Parity.None;
                RadioPort.StopBits = StopBits.One;
                RadioPort.RtsEnable = true;
                RadioPort.Open();
                RadioPortConnected = true;
                RadioPortComboBox.Enabled = false;
                CwRadioButton.Enabled = true;
                FmdRadioButton.Enabled = true;
                CwRadioButton.Checked = true;
                RadioBaudComboBox.Enabled = false;
                ChangeCW();
                RadioConnectButton.Text = "DisConnect";
            }
            else
            {
                RadioPort.Close();
                RadioPortConnected = false;
                RadioPortComboBox.Enabled = true;
                CwRadioButton.Enabled = false;
                FmdRadioButton.Enabled = false;
                RadioBaudComboBox.Enabled = true;
                RadioConnectButton.Text = "Connect";
            }

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
//            Properties.Settings.Default.Save();
            TncPort.Close();
            RadioPort.Close();
            RotatorPort.Close();
            trackingTimer.Stop();
        }

        private void CwRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            ChangeCW();
        }

        private void FmdRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            ChangeFMD();
        }

        //Downlink Terminal Setting
        private void DataReceived(object Sender, SerialDataReceivedEventArgs e)
        {
            AddRxDataDelegate add = new AddRxDataDelegate(AddRxData);
            byte[] SerialGetData = new byte[TncPort.BytesToRead];
            string ReceivedDataString = string.Empty;

            try
            {
                TncPort.Read(SerialGetData, 0, SerialGetData.Length);
            }

            catch
            {
                //Protect error
            }

            for (int i = 0; i < SerialGetData.Length; i++) {
                byte EachData = SerialGetData[i];
                ReceivedDataString += string.Format("{0:X2} ", EachData);
                if (EachData == 0xC0)
                {
                    if (ReceivedDataKissFrameFlag)
                    {
                        ReceivedDataKissFrameFlag = false;
                        ReceivedDataString += "\n";
                    }
                    else {
                        ReceivedDataKissFrameFlag = true;
                    }
                }
            }
            this.RxDataTextBox.Invoke(add, ReceivedDataString);

        }

        private delegate void AddRxDataDelegate(string data);
        private void AddRxData(string data)
        {

            this.RxDataTextBox.Focus();
            this.RxDataTextBox.AppendText(data);
            receivedPacketNumber += (data.Split('\n').Length - 1);
            receivedPacketsLabel.Text = "Received packet(s) : " + receivedPacketNumber;
        }


        private void ClearRxDataButton_Click(object sender, EventArgs e)
        {
            receivedPacketNumber = 0;
            receivedPacketsLabel.Text = "Received packet(s) : " + receivedPacketNumber;
            RxDataTextBox.Text = "";
        }
        private void SaveReceiveDataButton_Click(object sender, EventArgs e)
        {

            string rawRxData = RxDataTextBox.Text;

            DateTime dt = DateTime.UtcNow;
            string time = dt.ToString("yyyyMMdd_HHmmss");
            string folderPath = System.IO.Path.Combine(savingRootPath, satFolderName);
            string fileName;
            string absolutePath;

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }


            fileName = time + "raw.txt";
            absolutePath = System.IO.Path.Combine(folderPath, fileName);
            StreamWriter sw = new StreamWriter(absolutePath, false, Encoding.ASCII);
            sw.Write(rawRxData);
            sw.Close();

            fileName = dtAppStart.ToString("yyyyMMdd") + "log.txt";
            absolutePath = System.IO.Path.Combine(folderPath, fileName);
            sw = new StreamWriter(absolutePath, true, Encoding.ASCII);
            sw.Write(rawRxData);
            sw.Close();


            string[] rxDataRemovedComments = rawRxData.Split('#');
            string[] rxDataComments = new string[rxDataRemovedComments.Length];
            for (int i = 0; i < rxDataRemovedComments.Length; i++) {
                rxDataComments[i] = rxDataRemovedComments[i].Split('\n')[0];
                if (rxDataRemovedComments[i].Length > rxDataComments[i].Length)     //something contents are exist
                {
                    rxDataRemovedComments[i] = rxDataRemovedComments[i].Substring(rxDataComments[i].Length + 1);
                }
                else {
                    rxDataRemovedComments[i] = "";
                }
//                System.Diagnostics.Trace.WriteLine("comment" + i + ":" + rxDataComments[i]);
//                System.Diagnostics.Trace.WriteLine("content" + i + ":" + rxDataRemovedComments[i]);
//                System.Diagnostics.Trace.WriteLine("");
            }
            //2021/06/13 remove the function for compiling the packets and detecting the missing packets.
            ClearRxDataButton_Click(sender, e);
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
        }

        private void RxDataTextBox_SelectionChanged(object sender, EventArgs e)
        {
        }

        private void RxDataTextBox_TextChanged(object sender, EventArgs e)
        {
        }



        private const uint IMF_DUALFONT = 0x80;
        private const uint WM_USER = 0x0400;
        private const uint EM_SETLANGOPTIONS = WM_USER + 120;
        private const uint EM_GETLANGOPTIONS = WM_USER + 121;

        [System.Runtime.InteropServices.DllImport("USER32.dll")]
        private static extern uint SendMessage(
            System.IntPtr hWnd,
            uint msg,
            uint wParam,
            uint lParam);
        private void NoRichTextChange(RichTextBox RichTextBoxCtrl)
        {
            uint lParam;
            lParam = SendMessage(RichTextBoxCtrl.Handle, EM_GETLANGOPTIONS, 0, 0);
            lParam &= ~IMF_DUALFONT;
            SendMessage(RichTextBoxCtrl.Handle, EM_SETLANGOPTIONS, 0, lParam);
        }

        private void RxDataTextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
        }

        private void RxDataTextBox_VScroll(object sender, EventArgs e)
        {
        }

        private void Coments_Click(object sender, EventArgs e)
        {

        }

        private void Command_HEX_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {
        }

        private void SelectPathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            fbd.Description = "Select saving folder";
            fbd.SelectedPath = savingFolderPathTextBox.Text;
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                Console.WriteLine(fbd.SelectedPath);
                savingRootPath = fbd.SelectedPath;
                savingFolderPathTextBox.Text = savingRootPath;
                Properties.Settings.Default.savingPath = savingRootPath;
                Properties.Settings.Default.Save();
            }


        }

        private void ReceiveLabel_Click(object sender, EventArgs e)
        {

        }

        private void CommentsLabel_Click(object sender, EventArgs e)
        {

        }

        private void checkBoxAddMissinPacket_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void selectTlePathButton_Click(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "tle.txt";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "TLE file(*.txt)|*.txt|すべてのファイル(*.*)|*.*";
            //タイトルを設定する
            ofd.Title = "Select the TLE file";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(ofd.FileName);
                tlePath = ofd.FileName;
                tlePathTextBox.Text = tlePath;
                Properties.Settings.Default.tlePath = tlePath;
                Properties.Settings.Default.Save();
                setSatelliteComboList();
            }
        }

        private void latitudeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (latitudeTextBox.Text == "")
            {
                return;
            }
            try
            {
                latitude = Convert.ToDouble(latitudeTextBox.Text);
                Properties.Settings.Default.latitude = latitude;
                Properties.Settings.Default.Save();
            }
            catch
            {
                latitudeTextBox.Text = latitude.ToString("0.0000");
            }
        }

        private void longitudeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (longitudeTextBox.Text == "")
            {
                return;
            }
            try
            {
                longitude = Convert.ToDouble(longitudeTextBox.Text);
                Properties.Settings.Default.longitude = longitude;
                Properties.Settings.Default.Save();
            }
            catch
            {
                longitudeTextBox.Text = longitude.ToString("0.0000");
            }
        }

        private void heightTextBox_TextChanged(object sender, EventArgs e)
        {
            if (heightTextBox.Text == "")
            {
                return;
            }
            try
            {
                height = Convert.ToDouble(heightTextBox.Text);
                Properties.Settings.Default.height = Convert.ToDouble(heightTextBox.Text);
                Properties.Settings.Default.Save();
            }
            catch
            {
                heightTextBox.Text = height.ToString("0.0");
            }
        }

        private void SatelliteNameComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SatelliteNameComboBoxSub.SelectedIndex = SatelliteNameComboBox.SelectedIndex;
            Command_HEX.Text = "0000000000000000000000";
            Comments.Text = "None";
            SaveReceiveDataButton.Enabled = true;
            TransmitButton.Enabled = true;
        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void elevationLabel_Click(object sender, EventArgs e)
        {

        }

        private void civAddressTextBox_TextChanged(object sender, EventArgs e)
        {
            if (civAddressTextBox.Text == "") {
                return;
            }
            try
            {
                civAddress = Convert.ToInt32(civAddressTextBox.Text, 16);
                if (civAddress > 255) {
                    civAddress = Properties.Settings.Default.CIVaddress;
                    civAddressTextBox.Text = civAddress.ToString("X00");
                }
                Properties.Settings.Default.CIVaddress = civAddress;
                Properties.Settings.Default.Save();
            }
            catch {
                civAddressTextBox.Text = civAddress.ToString("X00");
            }

        }

        private void trackingCheckBox_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void RotatorConnectButton_Click(object sender, EventArgs e)
        {
            if (!RotatorPortConnected)
            {
                RotatorPort.PortName = RotatorPortComboBox.Text;
                RotatorPort.BaudRate = Convert.ToInt32(RotatorBaudComboBox.SelectedItem);
                RotatorPort.DataBits = 8;
                RotatorPort.Parity = Parity.None;
                RotatorPort.StopBits = StopBits.One;
                RotatorPort.RtsEnable = true;
                RotatorPort.Open();
                RotatorPortConnected = true;
                RotatorPortComboBox.Enabled = false;
                RotatorBaudComboBox.Enabled = false;
                RotatorConnectButton.Text = "DisConnect";
            }
            else
            {
                RotatorPort.Close();
                RotatorPortConnected = false;
                RotatorPortComboBox.Enabled = true;
                RotatorBaudComboBox.Enabled = true;
                RotatorConnectButton.Text = "Connect";
            }
        }

        private void frequencyAddjustValueLavel_Click(object sender, EventArgs e)
        {

        }

        private void buttonPulus500Hz_Click(object sender, EventArgs e)
        {
            freqAddjustValue += 500;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void buttonMinus500Hz_Click(object sender, EventArgs e)
        {
            freqAddjustValue -= 500;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void buttonPulus1kHz_Click(object sender, EventArgs e)
        {
            freqAddjustValue += 1000;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void buttonMinus1kHz_Click(object sender, EventArgs e)
        {
            freqAddjustValue -= 1000;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void buttonPulus5kHz_Click(object sender, EventArgs e)
        {
            freqAddjustValue += 5000;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void buttonMinus5kHz_Click(object sender, EventArgs e)
        {
            freqAddjustValue -= 5000;
            frequencyAddjustValueLavel.Text = freqAddjustValue.ToString();
        }

        private void AzHome_TextChanged(object sender, EventArgs e)
        {
            if (azHomeTextBox.Text == "")
            {
                return;
            }
            try
            {
                azHomeValue = Convert.ToDouble(azHomeTextBox.Text);
                Properties.Settings.Default.azHome = azHomeValue;
                Properties.Settings.Default.Save();
            }
            catch
            {
                azHomeTextBox.Text = azHomeValue.ToString("0.0");
            }

        }

        private void elHomeTextBox_TextChanged(object sender, EventArgs e)
        {
            if (elHomeTextBox.Text == "")
            {
                return;
            }
            try
            {
                elHomeValue = Convert.ToDouble(elHomeTextBox.Text);
                Properties.Settings.Default.elHome = elHomeValue;
                Properties.Settings.Default.Save();
            }
            catch
            {
                elHomeTextBox.Text = elHomeValue.ToString("0.0");
            }

        }

        private void rotatorModelComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.rotatorModel = rotatorModelComboBox.SelectedIndex;
            Properties.Settings.Default.Save();

        }

        private void RefleshButton_Click(object sender, EventArgs e)
        {
            if (!TncPortConnected && !RadioPortConnected && !RotatorPortConnected)
            {
                //Add port list into each COM port ComboBox
                string[] ports = SerialPort.GetPortNames();
                TncPortComboBox.Items.Clear();
                RadioPortComboBox.Items.Clear();
                RotatorPortComboBox.Items.Clear();
                foreach (string port in ports)
                {
                    TncPortComboBox.Items.Add(port);
                    RadioPortComboBox.Items.Add(port);
                    RotatorPortComboBox.Items.Add(port);
                }

                //Number of COM port is more than 1 -> TNC COM port is enabled
                if (TncPortComboBox.Items.Count >= 1)
                {
                    TncPortComboBox.SelectedIndex = 0;
                    //From BIRDS-1, for TNC Port, added Radio Port
                }

                //Number of COM port is more than 2 -> Radio COM port is enabled
                if (RadioPortComboBox.Items.Count >= 2)
                {
                    RadioPortComboBox.SelectedIndex = 1;
                }

                //Number of COM port is more than 3 -> Rotator COM port is enabled
                if (RotatorPortComboBox.Items.Count >= 3)
                {
                    RotatorPortComboBox.SelectedIndex = 2;
                }

            }
            else {
                MessageBox.Show("Please disconnect all com port");
            }
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void SatelliteNameComboBoxSub_SelectedIndexChanged(object sender, EventArgs e)
        {
            SatelliteNameComboBox.SelectedIndex = SatelliteNameComboBoxSub.SelectedIndex;

            using (var parser = new TextFieldParser(freqPath))
            {
                parser.Delimiters = new string[] { "," };
                while (!parser.EndOfData)
                {
                    var fields = parser.ReadFields();
                    if (fields[0] == SatelliteNameComboBoxSub.SelectedItem.ToString())
                    {
                        noradIdLabel.Text = fields[1];

                        cwFreq = Convert.ToDouble(fields[2]);
                        ulFreq = Convert.ToDouble(fields[3]);
                        dlFreq = Convert.ToDouble(fields[4]);


                        dlFreqLabel.Text = (dlFreq/1000000).ToString();
                        ulFreqLabel.Text = (ulFreq/1000000).ToString();
                        remarkLabel.Text = fields[5];
                        folderNameLabel.Text = fields[6];
                        satFolderName = fields[6];

                        foreach (var tle in tleList)
                        {
                            if (tle.getNoradID() == noradIdLabel.Text) {
                                selectedTle = tle;
                                break;
                            }
                        }

                        break;
                    }
                }
            }


        }

        private void button1_Click(object sender, EventArgs e)  //Freq List Button
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "FreqList.csv";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "TLE file(*.csv)|*.csv|すべてのファイル(*.*)|*.*";
            //タイトルを設定する
            ofd.Title = "Select the frequency list file";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(ofd.FileName);
                freqPath = ofd.FileName;
                freqPathTextBox.Text = freqPath;
                Properties.Settings.Default.freqPath = freqPath;
                Properties.Settings.Default.Save();
                setSatelliteComboList();
            }

        }

        private void latitudeTextBox_Enter(object sender, EventArgs e)
        {

        }
    }
}
