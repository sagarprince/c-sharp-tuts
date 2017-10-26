using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Speech.Synthesis;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Web;

namespace JARVIS_PROJECT
{
    public partial class Form2 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine();
        SpeechSynthesizer JARVIS = new SpeechSynthesizer();
        string Temperature;
        string Condition;
        string Humidity;
        string WindSpeed;
        string Town;
        string TFCond;
        string TFHigh;
        string TFLow;
        string BrowseDirectory;
        string QEvent;
        string ProcWindow;
        double timer = 10;
        int count = 1;
        Random rnd = new Random();

        public Form2()
        {
            InitializeComponent();

            _recognizer.SetInputToDefaultAudioDevice();
            _recognizer.LoadGrammar(new DictationGrammar());
            _recognizer.LoadGrammar(new Grammar(new GrammarBuilder(new Choices(System.IO.File.ReadAllLines(@"C:\Users\user\Documents\A.L.E.X.\Commands.txt")))));
            _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(_recognizer_SpeechRecognized);
            _recognizer.RecognizeAsync(RecognizeMode.Multiple);

        }
        enum RecycleFlags : uint
        {
            SHERB_NOCONFIRMATION = 0x00000001,
            SHERB_NOPROGRESSUI = 0x00000001,
            SHERB_NOSOUND = 0x00000004
        }

        [DllImport("Shell32.dll", CharSet = CharSet.Unicode)]
        static extern uint SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath,
        RecycleFlags dwFlags);


        private void StopWindow()
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(ProcWindow);
            foreach (System.Diagnostics.Process proc in procs)
            {
                proc.CloseMainWindow();
            }
        }


        private void ComputerTermination()
        {
            switch (QEvent)
            {
                case "shutdown":
                    System.Diagnostics.Process.Start("shutdown", "-s");
                    break;
                case "logoff":
                    System.Diagnostics.Process.Start("shutdown", "-l");
                    break;
                case "restart":
                    System.Diagnostics.Process.Start("shutdown", "-r");
                    break;
            }
        }
        private void GetWeather()
        {

            string query = String.Format("http://weather.yahooapis.com/forecastrss?w=12807773");
            XmlDocument wData = new XmlDocument();
            wData.Load(query);

            XmlNamespaceManager manager = new XmlNamespaceManager(wData.NameTable);
            manager.AddNamespace("yweather", "http://xml.weather.yahoo.com/ns/rss/1.0");

            XmlNode channel = wData.SelectSingleNode("rss").SelectSingleNode("channel");
            XmlNodeList nodes = wData.SelectNodes("/rss/channel/item/yweather:forecast", manager);

            Temperature = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["temp"].Value;
            Condition = channel.SelectSingleNode("item").SelectSingleNode("yweather:condition", manager).Attributes["text"].Value;
            Humidity = channel.SelectSingleNode("yweather:atmosphere", manager).Attributes["humidity"].Value;
            WindSpeed = channel.SelectSingleNode("yweather:wind", manager).Attributes["speed"].Value;
            Town = channel.SelectSingleNode("yweather:location", manager).Attributes["city"].Value;
            TFCond = channel.SelectSingleNode("item", manager).SelectSingleNode("yweather:forecast", manager).Attributes["text"].Value;
            TFHigh = channel.SelectSingleNode("item", manager).SelectSingleNode("yweather:forecast", manager).Attributes["high"].Value;
            TFLow = channel.SelectSingleNode("item", manager).SelectSingleNode("yweather:forecast", manager).Attributes["low"].Value;
        }

        void _recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            int ranNum = rnd.Next(1, 10);
            string speech = e.Result.Text;
            switch (speech)
            {

                //DATE,DAY,TIME,WEATHER
                case "Hows the weather":
                    GetWeather();
                    JARVIS.Speak("The weather in" + Town + "is" + Condition + "at" + Temperature + "degrees. There is a wind speed of" + WindSpeed + "miles per hour and humidity of" + Humidity);
                    break;
                case "Whats tomorrows forecast":
                    GetWeather();
                    JARVIS.Speak("It looks like tomorrow will be" + TFCond + "with the high of" + TFHigh + "and a low of" + TFLow);
                    break;
                case "What time is it":
                case "Whats the time":
                case "What time it is":
                    DateTime now = DateTime.Now;
                    string time = now.GetDateTimeFormats('t')[0];
                    JARVIS.Speak(time);
                    break;
                case "What day is it":
                    JARVIS.Speak(DateTime.Today.ToString("dddd"));
                    break;
                case "Whats the date":
                case "Whats todays date":
                    JARVIS.Speak(DateTime.Today.ToString("dd-MM-yyyy"));
                    break;


                //MAXIMIZE AND MINIMIZE
                case "Out of the way":
                case "Move":
                case "Minimize":
                    JARVIS.Speak("I apologize sir");
                    WindowState = FormWindowState.Minimized;
                    break;
                case "Go back":
                case "Come back":
                case "Return":
                    JARVIS.Speak("Right away sir!");
                    WindowState = FormWindowState.Normal;
                    break;

                // UTILITY
                case "Chrome":
                    JARVIS.Speak("Right away sir!");
                    Process.Start(@"C:\Users\user\AppData\Local\Google\Chrome\Application\chrome.exe");
                    break;
                case "Firefox":
                    JARVIS.Speak("Little less favourite, but still usable sir!");
                    Process.Start(@"C:\Program Files\Mozilla Firefox\firefox.exe");
                    break;
                case "Facebook":
                    JARVIS.Speak("What a dumb choice, but here you go! ");
                    Process.Start(@"www.facebook.com");
                    break;
                case "Youtube":
                case "Music":
                    JARVIS.Speak("We could use some music sir! Don't get too excited!");
                    Process.Start(@"www.youtube.com");
                    break;
                case "Wikipedia":
                case "Wiki":
                    JARVIS.Speak("Finally something smart sir!");
                    Process.Start(@"www.wikipedia.com");
                    break;
                case "Check my mail":
                    JARVIS.Speak("Checking your mail sir");
                    Process.Start(@"www.gmail.com");
                    break;
                case "Check mail":
                    JARVIS.Speak("Checking your mail sir");
                    Process.Start(@"www.outlook.com");
                    break;


                // GREETINGS
                case "Hello":
                case "Hello Alex":
                    JARVIS.Speak("Hello sir!");
                    break;

                case "Goodbye":
                case "Goodbye Alex":
                case "Farewell":
                    JARVIS.Speak("Until next time sir!");
                    Application.Exit();
                    break;

                // Alex
                case "Alex":
                    if (ranNum < 6) { JARVIS.Speak("Hello sir"); }
                    else if (ranNum > 5) { JARVIS.Speak("Hi"); }
                    break;
                case "A":
                    if (ranNum < 6) { JARVIS.Speak("Yo sir!"); }
                    else if (ranNum > 5) { JARVIS.Speak("Welcome back sir!"); }
                    break;
                case "Wake up":
                    if (ranNum < 6) { JARVIS.Speak("I'm up sir! What can I do for you!"); }
                    else if (ranNum > 5) { JARVIS.Speak("I'm here sir!"); }
                    break;
                case "Wake up daddys home":
                    if (ranNum < 6) { JARVIS.Speak("How lovely to see you again sir"); }
                    else if (ranNum > 5) { JARVIS.Speak("Hello sir! What can I do for yo today"); }
                    break;

                case "Thanks":
                    JARVIS.Speak("You're welcome sir!");
                    break;

                case "Thank you":
                    JARVIS.Speak("Don't mention it sir! I'll always be here!");
                    break;

                // EMPTY RECYCLE BIN
                case "Empty Recycle Bin":
                    if (ranNum < 6) { JARVIS.Speak("Right away sir!"); }
                    else if (ranNum > 5) { JARVIS.Speak("As you wish sir!"); }
                    uint result = SHEmptyRecycleBin(IntPtr.Zero, null, 0);
                    break;


                // SHUTDOWN RESTART LOGOFF
                case "Jarvis log off":
                    new Login().Show();
                    this.Hide();
                    break;
                case "Shutdown":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "Shutdown";
                        JARVIS.Speak("I will shutdown shortly");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Log off":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "Logoff";
                        JARVIS.Speak("Logging off");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Restart":
                    if (ShutdownTimer.Enabled == false)
                    {
                        QEvent = "Restart";
                        JARVIS.Speak("I'll be back shortly");
                        lblTimer.Visible = true;
                        ShutdownTimer.Enabled = true;
                    }
                    break;
                case "Abort":
                case "Cancel":
                    if (ShutdownTimer.Enabled == true)
                    {
                        QEvent = "Abort";
                        QEvent = "Cancel";
                    }
                    break;
                case "Speed up":
                    if (ShutdownTimer.Enabled == true)
                    {
                        ShutdownTimer.Interval = ShutdownTimer.Interval / 10;
                    }
                    break;
                case "Slow down":
                    if (ShutdownTimer.Enabled == true)
                    {
                        ShutdownTimer.Interval = ShutdownTimer.Interval * 10;
                    }
                    break;

                // LOAD DIRECTORY
                case "load music":
                    QEvent = "music";
                    JARVIS.Speak("In a mood for some music. Excelent");
                    BrowseDirectory = @"C:\Users\user\Music\";
                    LoadDirectory();
                    listBox1.Visible = true;
                    break;
                case "load pictures":
                    QEvent = "pictures";
                    JARVIS.Speak("Right away sir");
                    BrowseDirectory = @"C:\Users\user\Pictures\";
                    LoadDirectory();
                    listBox1.Visible = true;
                    break;
                case "load videos":
                    QEvent = "videos";
                    JARVIS.Speak("Enjoy sir!");
                    BrowseDirectory = @"C:\Users\user\Videos\";
                    LoadDirectory();
                    listBox1.Visible = true;
                    break;
                case "browse":
                    FolderBrowserDialog fbd = new FolderBrowserDialog();
                    if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        BrowseDirectory = fbd.SelectedPath;
                    }
                    QEvent = "load directory";
                    LoadDirectory();
                    listBox1.Visible = true;
                    break;
                // MUSIC AND MANY OTHERS
                case "Play some music":
                    JARVIS.Speak("Do you want to choose between genre or band");
                    break;
                case "Genre":
                    JARVIS.Speak("I'm in a mood for some music too. Choose genre: rock, hip-hop, rap, r'n'b etc.");
                    break;
                case "Band":
                    JARVIS.Speak("Choose a band sir! You have: Bullet For My Valentine; Nickleback; Three Days Grace and many other, but I believe those three are your favourite");
                    break;
                case "Bullet For My Valentine":
                    QEvent = "bulletformyvalentine";
                    JARVIS.Speak("You sure!");
                    break;
                case "Nickelback":
                    QEvent = "nickelback";
                    JARVIS.Speak("Are you sure");
                    break;
                case "Three Days Grace":
                    QEvent = "threedaysgrace";
                    JARVIS.Speak("Speak Yes and we will go");
                    break;
                case "Black Veil Brides":
                    QEvent = "bvb";
                    JARVIS.Speak("Are you positive sir!");
                    break;
                case "Hip-Hop":
                    QEvent = "HipHop";
                    JARVIS.Speak("Are you sure you want hip-hop music sir!");
                    break;
                case "Rock":
                    QEvent = "Rock";
                    JARVIS.Speak("Are you sure you want rock. Feeling in a little destruction mood are we. Very well sir, if you like so will I");
                    break;
                case "Rap":
                    QEvent = "Rap";
                    JARVIS.Speak("Do you really want rap sir!");
                    break;
                case "RnB":
                    QEvent = "RnB";
                    JARVIS.Speak("You sure sir!");
                    break;
                case "Metal":
                    QEvent = "Metal";
                    JARVIS.Speak("Are you sure sir!");
                    break;

                case "Yes":
                case "OK":
                case "Alright":
                    if (QEvent == "Hip-Hop")
                    {
                        Process.Start(@"http://youtu.be/PZca3ZZEptI");
                    }
                    else if (QEvent == "Rock")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=ioFMHbkntVU&feature=share&list=PL39BF36EE68E9FE01");
                    }
                    else if (QEvent == "Rap")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=uelHwf8o7_U&list=PLE0D603BE945F8E59");
                    }
                    else if (QEvent == "RnB")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=dGghkjpNCQ8&feature=share&list=PL3770D578D4B71793");
                    }
                    else if (QEvent == "Metal")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=CEb3GQqdHD4&feature=share&list=PLAA46E0AFDFABFC7D");
                    }
                    else if (QEvent == "bulletformyvalentine")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=ioFMHbkntVU&feature=share&list=PL39BF36EE68E9FE01");
                    }
                    else if (QEvent == "nickelback")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=1cQh1ccqu8M&feature=share&list=PLAC97E191247EA0BA");
                    }
                    else if (QEvent == "threedaysgrace")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=vDTK9Rrg03w&feature=share&list=PL0A6A2F0BED1B2312");
                    }
                    else if (QEvent == "bvb")
                    {
                        Process.Start(@"http://www.youtube.com/watch?v=CEb3GQqdHD4&feature=share&list=PLAA46E0AFDFABFC7D");
                    }
                    break;

            }
        }

        private void procedure()
        {
            throw new NotImplementedException();
        }

private void ShutdownTimer_Tick(object sender, EventArgs e)
        {
            if (timer == 0)
            {
                lblTimer.Visible = false;
                ComputerTermination();
                ShutdownTimer.Enabled = false;
            }
            else if (QEvent == "Abort" || QEvent == "Cancel")
            {
                timer = 10;
                lblTimer.Visible = false;
                ShutdownTimer.Enabled = false;
            }
            else
            {
                timer = timer - .01;
                lblTimer.Text = timer.ToString();
            }
        }

        private void LoadDirectory()
        {
            listBox1.Items.Clear();
            switch (QEvent)
            {
                case "music":
                    string[] files = Directory.GetFiles(BrowseDirectory, "*.mp3", SearchOption.AllDirectories);
                    foreach (string file in files) { listBox1.Items.Add(file.Replace(BrowseDirectory, "")); }
                    break;

                case "pictures":
                    files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in files) { listBox1.Items.Add(file.Replace(BrowseDirectory, "")); }
                    break;

                case "videos":
                    files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in files) { listBox1.Items.Add(file.Replace(BrowseDirectory, "")); }
                    break;

                case "browse":
                    files = Directory.GetFiles(BrowseDirectory, "*", SearchOption.AllDirectories);
                    foreach (string file in files) { listBox1.Items.Add(file.Replace(BrowseDirectory, "")); }
                    break;
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Object open = BrowseDirectory + listBox1.SelectedItem;
            try
            {
                Process.Start(open.ToString());
            }
            catch
            {
                open = BrowseDirectory + "\\" + listBox1.SelectedItem; Process.Start(open.ToString());
            }
        }


Enjoy.