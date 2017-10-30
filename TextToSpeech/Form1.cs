using System;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Threading;
using System.Speech.Synthesis;
using System.Speech.Recognition;

namespace TextToSpeech
{
    public partial class Form1 : Form
    {
        bool cikisDurum = false;
        DateTime time = new DateTime();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        RegistryKey key = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Policies\\System");

        string[] konusulan;
        int kelimesayaci = 0;
        public Form1()
        {
            InitializeComponent();
            timer1.Start();
            key.SetValue("DisableTaskMgr", 0);
            time = DateTime.Now;
            SubTitle.Text = "";

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!cikisDurum)
            {
                e.Cancel = true;
                Konus("You can not close me.");
            }
            else
                e.Cancel = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!cikisDurum)
            {
                SetForegroundWindow(this.Handle);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cikisDurum = true;
            Application.Exit();
        }

        SpeechSynthesizer sSynth = new SpeechSynthesizer();
        PromptBuilder pBuilder = new PromptBuilder();
        SpeechRecognitionEngine sRecognize = new SpeechRecognitionEngine();

        void Konus(string metin)
        {
            SubTitle.Text = metin;
            sSynth.SelectVoice("Microsoft David Desktop");
            pBuilder.ClearContent();
            pBuilder.AppendText(metin);
            sSynth.Speak(pBuilder);
            SubTitle.Text = "";
        }

        void AltYazi(string metin)
        {
            SubTitle.Visible = true;
            konusulan = metin.Split();

            foreach (var item in konusulan)
            {
                kelimesayaci += 1;
            }
            SubTitleTimer.Start();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            Konus("Welcome back sir!");
            Konus("open");
            // Konus("How are you?");
        }


        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            TimeSpan fark = DateTime.Now - time;
            if (fark.Minutes > 10)
                Konus("Welcome Back Sir");
            time = DateTime.Now;

            if (e.KeyCode == Keys.Space)
            {
                if (!cikisDurum)
                    cikisDurum = true;
                else
                    cikisDurum = false;
            }
            

            if (e.KeyCode == Keys.ControlKey)
            {
                if (!cikisDurum)
                {
                    Konus("You can not close me.");
                }

            }
            if (e.KeyCode == Keys.T)
            {
                Konus("I am listening sir!");
                Choices sList = new Choices();
                sList.Add(new string[] { "start chrome", "hello", "test", "it works", "how", "are", "you", "today", "i", "am", "fine", "exit", "so" });
                Grammar gr = new Grammar(new GrammarBuilder(sList));
                try
                {
                    sRecognize.RequestRecognizerUpdate();
                    sRecognize.LoadGrammar(gr);
                    sRecognize.SpeechRecognized += sRecognize_SpeechRecognized;
                    sRecognize.SetInputToDefaultAudioDevice();
                    sRecognize.RecognizeAsync(RecognizeMode.Multiple);

                }
                catch (Exception)
                {

                    return;
                }
            }
        }
        public string sonKelime = "";
        private void sRecognize_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text == "exit")
            {
                cikisDurum = true;
                Application.Exit();
            }
            else if (e.Result.Text == "start chrome"&& sonKelime!="start chrome")
            {
                //MessageBox.Show("ok");
                System.Diagnostics.Process.Start("chrome.exe", "http://www.google.com.tr");

            }  
            //MessageBox.Show("speech recognized:" + e.Result.Text.ToString());
            sonKelime = e.Result.Text.ToString();
        }
        string[] elemansil(string[] dizi)
        {
            string[] dizi2 = new string[dizi.Length - 1];
            int sayac = dizi.Length;
            for (int i = 1; i < sayac; i++)
            {
                dizi2[i - 1] = dizi[i];
            }
            return dizi2;
        }
        private void SubTitleTimer_Tick(object sender, EventArgs e)
        {

        }
    }
}
