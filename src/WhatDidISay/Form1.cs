using System;
using System.Windows.Forms;
using NAudio.Wave;
using NAudio.Utils;
using System.IO;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices; //required for dll import

namespace WhatDidISay
{
    internal partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifers, int vlc);

        public Form1()
        {
            // CTRL-ALT-SHIFT-X
            RegisterHotKey(this.Handle, HotkeyId, 7, (int)Keys.X);

            InitializeComponent();
        }

        IWaveIn _wavein;
        // cache number of inputs because it can change during running this program
        private int inputCount;
        private CircularBuffer _circularBuffer;
        private byte[] _dataBuffer;
        internal Options Options { get; set; }
        private const int HotkeyId = 1;

        private void OnStartBuffering_Click(object sender, EventArgs e)
        {
            _wavein.DataAvailable += OnReceiveAudio;
            WindowState = FormWindowState.Minimized;
            ntfyIcon.BalloonTipText = string.Format(Resource.BufferingMinutesOfAudio, txtBufferSize);
            ntfyIcon.Visible = true;
        }

        void OnReceiveAudio(object sender, WaveInEventArgs e)
        {
            _circularBuffer.Write(e.Buffer, 0, e.BytesRecorded);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            txtBufferSize.Text = Options.BufferTimeInMinutes.ToString();
            cmbRecordingDevices.Items.Clear();
            cmbRecordingDevices.Enabled = false;
            btnStartBuffering.Enabled = false;
            inputCount = WaveIn.DeviceCount;

            for (int device = 0; device < inputCount; device++)
            {
                var capabilities = WaveIn.GetCapabilities(device);
                cmbRecordingDevices.Items.Add(capabilities.ProductName);
            }
            cmbRecordingDevices.Items.Add("System output");

            cmbRecordingDevices.SelectedIndex = 0;
            cmbRecordingDevices.Enabled = true;
            btnStartBuffering.Enabled = true;

            CreateWaveProvider();
        }
        
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0312 && m.WParam.ToInt32() == HotkeyId)
            {
                MouseEventArgs e = new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0);
                OnNotifyIcon_Click(null, e);
            }
            base.WndProc(ref m);
        }

        private void OnComboDevices_Select(object sender, EventArgs e)
        {
            CreateWaveProvider();
        }

        private void CreateWaveProvider()
        {
            var bufferTimeInSeconds = int.Parse(txtBufferSize.Text) * 60;
            var bufferSizeInBytes = Options.SampleRate * (Options.BitsPerSample / 8) * Options.Channels * bufferTimeInSeconds;
            _circularBuffer = new CircularBuffer(bufferSizeInBytes);
            _dataBuffer = new byte[bufferSizeInBytes];

            _wavein?.Dispose();

            var cmbRecordingDevices_selected = cmbRecordingDevices.SelectedIndex;
            if (cmbRecordingDevices_selected < inputCount)
            {
                _wavein = new WaveIn
                {
                    WaveFormat = new WaveFormat(Options.SampleRate, Options.BitsPerSample, Options.Channels),
                    BufferMilliseconds = 500,
                    NumberOfBuffers = 3,
                    DeviceNumber = cmbRecordingDevices.SelectedIndex,
                };
            }
            else
            {
                // System output is selected
                // TODO: if we change output during recording (like switching on headset which triggers switching from speakers to headset), it doesn't follow
                _wavein = new WasapiLoopbackCapture
                {
                    WaveFormat = new WaveFormat(Options.SampleRate, Options.BitsPerSample, Options.Channels),
                };
            }

            _wavein.StartRecording();
            _wavein.DataAvailable += VuMeter;
        }

        private void VuMeter(object sender, WaveInEventArgs e)
        {            
            short maxvu = 0;
            for (int i = 0; i < e.BytesRecorded; i+=4)
            {
                var vu = (short)(e.Buffer[i + 1] *256 + e.Buffer[i]);
                if (vu > maxvu) maxvu = vu;
            }

            var vuValue = 96 + (int)Math.Max(20.0f * Math.Log10(maxvu / 32768.0f), -96.0f);
            SetVuMeter(vuValue);
        }

        // to access vumeter.Value in a thread-safe way. In case of recording system output, VuMeter() callback is called from an other thread.
        private delegate void SetVuMeterCallback(int val);

        private void SetVuMeter(int val)
        {
            if (vumeter.InvokeRequired)
            {
                SetVuMeterCallback d = new SetVuMeterCallback(SetVuMeter);
                vumeter.Invoke(d, new object[] { val });
            }
            else
            {
                vumeter.Value = val;
            }
        }

        private string getFileName()
        {
            var fileCnt = 1;
            var outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "WhatDidISay");
            if (Directory.Exists(outputFolder))
            {
                Regex regex = new Regex(@"\brecorded_(\d+)\.wav$");
                foreach (string f in Directory.EnumerateFiles(outputFolder))
                {
                    Match match = regex.Match(f);
                    if (match.Success)
                    {
                        var newFileCnt = int.Parse(match.Groups[1].Value);
                        fileCnt = Math.Max(fileCnt, newFileCnt + 1);
                    }
                }
            }
            else
            {
                Directory.CreateDirectory(outputFolder);
            }

            return Path.Combine(outputFolder, "recorded_" + fileCnt.ToString() + ".wav");
        }

        private void OnNotifyIcon_Click(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ntfyIcon.Visible = false;
                WindowState = FormWindowState.Normal;
                ShowInTaskbar = true;
                return;
            }

            var currentBufferSize = _circularBuffer.Count;
            _circularBuffer.Read(_dataBuffer, 0, currentBufferSize);
            
            using (var wfw = new WaveFileWriter(getFileName(), _wavein.WaveFormat))
            {
                wfw.Write(_dataBuffer, 0, currentBufferSize);
            }
        }
    }
}