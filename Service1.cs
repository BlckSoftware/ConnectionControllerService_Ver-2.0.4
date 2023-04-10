using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Configuration;
using System.Threading.Tasks;
using System.Timers;
using System.Net;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Globalization;


namespace ConnectionControllerService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer(); // name space(using System.Timers;)        
       
        DataTable table = new DataTable();
        public string MachineName = Environment.MachineName;
        public Service1()
        {
            InitializeComponent();
        }
       public int Min = Convert.ToInt32(ConfigurationManager.AppSettings["Min"]);     
       
        #region servis ayarı
        protected override void OnStart(string[] args)
        {            
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture = new CultureInfo("tr-TR"));
            checkFiles();
            data();
            try { 
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = Min*60000; //number in milisecinds  
            timer.Enabled = true;
            timer.AutoReset = true;       
            WriteToFile(zaman + " | ********** SERVİS BAŞLATILDI ********** " );        
                }
            catch(Exception ex) { WriteToFile("SERVİS BAŞLATILAMADI. EXCEPTION :" + ex.Message); }
            ICMP();            
        }
        protected override void OnStop()
        {
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture = new CultureInfo("tr-TR"));
            WriteToFile(zaman + " | ********** SERVİS DURDURULDU ********** ");
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + "mail.txt"))
               {File.Delete(AppDomain.CurrentDomain.BaseDirectory + "mail.txt");}
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            ICMP();  
        }
        #endregion
       
        #region ICMP Fonksiyonu
        public void ICMP()
        {
            string min = ConfigurationManager.AppSettings["Min"];
            string z = ConfigurationManager.AppSettings["Key"];
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture);
            WriteToFile(zaman + " | ********** BAGLANTILAR KONTROL EDİLİYOR **********");
            try { 
            var ip_list = table.AsEnumerable().Select(s => s.Field<string>("CİHAZ İP"));          
                foreach (var k in ip_list)
                {
                    Ping pingSender = new Ping();
                    //  var numbers = new[] { File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + "Cihaz_Listesi.txt") };
                    // var words = new[] { File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + "Cihaz_Listesi.txt") };
                    int timeout = 5000;// Wait 5 seconds for a reply.
                    // Create a buffer of 32 bytes of data to be transmitted.
                    string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
                    byte[] buffer = Encoding.ASCII.GetBytes(data);
                    PingOptions options = new PingOptions(64, true);
                    PingReply reply = pingSender.Send(k, timeout, buffer, options); // Send the request.

                    if (reply.Status == IPStatus.Success)
                    {
                    
                    DataRow[] ad = table.Select("[CİHAZ İP] = '" + k + "'");
                        foreach (DataRow row in ad)
                        {
                            string cihaz = (row[0] + " | " + row[1]);
                            string logmes_bas = zaman + " | İstasyon : " + cihaz + "  Bağlantı Var. ";
                            WriteToFile(logmes_bas);
                        }
                    }
                    else if (reply.Status != IPStatus.Success)
                    {
                    
                    DataRow[] adı = table.Select("[CİHAZ İP] = '" + k + "'");

                    foreach (DataRow row in adı)
                    {
                        string cihaz = (row[0] + " | " + row[1]);
                        string logmes = zaman + " | İstasyon : " + cihaz + "   Bağlantı Yok. ";
                        WriteToFile(logmes);
                        if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "mail.txt"))
                        {
                            File.Create(AppDomain.CurrentDomain.BaseDirectory + "mail.txt").Close();
                        }
                        using (StreamWriter sw = File.AppendText(AppDomain.CurrentDomain.BaseDirectory + "mail.txt"))// log dosyasına mesaj yazma.
                        {
                            sw.WriteLine(logmes); // içerisine logmes yaz.
                        }
                    }                         
                }                
             }
          }
            catch (Exception ex) { WriteToFile("ICMP Fonksiyonu Başlatılamadı. \n Exception : "+ex.Message); }
            if (z == "1"&& File.Exists(AppDomain.CurrentDomain.BaseDirectory + "mail.txt")) 
            {
               try {                
                
                Yandex(); }
                catch(Exception ex) { WriteToFile(zaman+" | !!!!MAIL GÖNDERİLEMEDİ!!! |"+ex.Message); }
                }
            else { return; }          
            
        }
        #endregion
        #region Log yazdırma
        public void WriteToFile(string Message)
         {
             string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
             if (!Directory.Exists(path))
             {
                 Directory.CreateDirectory(path);
             }
             string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
             if (!File.Exists(filepath))
             {
                 // Create a file to write to.   
                 using (StreamWriter sw = File.CreateText(filepath))
                 {
                     sw.WriteLine(Message);
                 }
             }
             else
             {
                 using (StreamWriter sw = File.AppendText(filepath))
                 {
                     sw.WriteLine(Message);
                 }

             }
         }
        #endregion
        #region Dosya Kontrolü
        public void checkFiles()
         {
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture = new CultureInfo("tr-TR"));
            if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory +AppDomain.CurrentDomain.FriendlyName+".config"))
            {
                string message = ("Config Dosyası Bulunamadı. Servis Başlatılamadı. \n" + AppDomain.CurrentDomain.FriendlyName + ".config" + "  Adında xml config dosyası gerekmektedir. \n huseyinkarayazim@gmail.com 'ile iletişime geçiniz. \n\n ");
                WriteToFile(message);
            }
             if (!File.Exists(AppDomain.CurrentDomain.BaseDirectory + "Cihaz_Listesi.txt"))
             {
                 File.Create(AppDomain.CurrentDomain.BaseDirectory + "Cihaz_Listesi.txt").Close();
                 string message = "''Cihaz_Listesi.txt BULUNAMADI !!''\n''Cihaz_Listesi'' ADINDA TEXT DOSYASI OLUŞTURULDU\n " + AppDomain.CurrentDomain.BaseDirectory + "\nREADME.TXT GÖZ ATIN";
                 WriteToFile(message + zaman);
             }
        }
        #endregion
        #region data okuma
        public void data()
         {
            string sep =ConfigurationManager.AppSettings["Separator"];
            char[] c = sep.ToCharArray();
            table.Columns.Add("CİHAZ ADI", typeof(string));    // CIHAZ_ADI adında sütun oluşturuldu .
             table.Columns.Add("CİHAZ İP", typeof(string));     // CIHAZ_IP adında sütun oluşturuldu .
            try { 
            string[] liste = File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + "Cihaz_Listesi.txt"); // Tablo için text dosyası okundu.
            string []values;  // for için değişken tanımlandı.
            for (int i = 0; i < liste.Length; i++) // tablodaki değerler kadar for döngüsü başlatıldı.
             { values = liste[i].ToString().Split(c);   // tabloda "/" ile sütunlar ayrıldı .
                 string[] row = new string[values.Length];   // sütunları okumak için for döngüsü yapıldı.
                 for (int j = 0; j < values.Length; j++)
                 {
                     row[j] = values[j].Trim();
                 }
                 table.Rows.Add(row);  // Text deki değerler tabloya aktarıldı.
             }
            }
            catch (Exception ex) { WriteToFile("Cihaz_Listesi.txt klasörü boş.\n\n"+ex.Message); }
        }
        #endregion
        #region gmail ile mail gönderme 
        public void send_mail()
         {
            string email = ConfigurationManager.AppSettings["EMAIL"];
            string pass= ConfigurationManager.AppSettings["PASS"];
            int port=Convert.ToInt32(ConfigurationManager.AppSettings["PORT"]);
            string host= ConfigurationManager.AppSettings["HOST"];
            string alıcılar = ConfigurationManager.AppSettings["To"];
            string tesis = ConfigurationManager.AppSettings["TESIS_ADI"];
            string subject = (tesis + " " + " HABERLEŞME UYARI SERVİSİ ");
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture);
            var log = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "mail.txt");
             MailMessage error = new MailMessage();
             SmtpClient istemci = new SmtpClient();
             istemci.Port =port;
             istemci.Host = host;
             istemci.EnableSsl = true;
             istemci.Credentials = new System.Net.NetworkCredential(email, pass);
             error.To.Add(alıcılar);
             error.From = new MailAddress(email);
             error.Subject = (subject);
             error.Body = (" \n \n DİKKAT\n\n Sayın Yetkili Aşağıdaki İstasyon / İstasyonlara Bağlantı Sağlanamamaktadır.\n\n \n " + log + " \n\n\n Bu Mail Connection Controller Service İle "+MachineName+" Tarafından  Gönderildi \n\n\n Bir Sonraki Kontrol "+Min+"   Dakika Sonra \n \n \n Tarih : " + zaman+"\n \n \n Daha Fazla Bilgi İçin "+MachineName+"':\'"+ AppDomain.CurrentDomain.BaseDirectory + "README.TXT GÖZ ATIN"+"\n \n Yada 'huseyinkarayazim@gmail.com' ile iletişime geçin. ");
             istemci.Send(error);
            File.Delete(AppDomain.CurrentDomain.BaseDirectory + "mail.txt");
          }
        #endregion       
        #region Yandex ile mail gönderme
        public void Yandex()
        {          
            string zaman = DateTime.Now.ToString("f", CultureInfo.CurrentCulture = new CultureInfo("tr-TR"));
            var log = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + "mail.txt");
            if (log != null)
            {
                string alıcılar = ConfigurationManager.AppSettings["To"];
                string tesis = ConfigurationManager.AppSettings["TESIS_ADI"];
                bool SSL =Convert.ToBoolean(ConfigurationManager.AppSettings["SSL"]);
                string subject = (tesis + " " + " HABERLEŞME UYARI SERVİSİ ");
                string content = (" \n \n DİKKAT \n\n " + tesis + "\n\n Sayın Yetkili Aşağıdaki İstasyon / İstasyonlara Bağlantı Sağlanamamaktadır.\n\n \n " 
                    + log + " \n\n\n Bu Mail Connection Controller Service İle " + MachineName + " Tarafından  Gönderildi \n\n\n Bir Sonraki Kontrol " +Min+ " Dakika Sonra \n \n \n Tarih : " + zaman + "\n \n \n Daha Fazla Bilgi İçin " + MachineName + "':\'" + AppDomain.CurrentDomain.BaseDirectory + "README.TXT GÖZ ATIN" + "\n \n Yada 'huseyinkarayazim@gmail.com / huseyin.karayazim@triatech.com.tr' ile iletişime geçin.\n\n ");
                var _host = ConfigurationManager.AppSettings["HOST"];
                var _port = Convert.ToInt32(ConfigurationManager.AppSettings["PORT"]);
                var _defaultCredentials = false;
                var _enableSsl = SSL;
                var _emailfrom = ConfigurationManager.AppSettings["EMAIL"];//Your yandex email adress
                var _password = ConfigurationManager.AppSettings["PASS"];//Your yandex app password
                using (var smtpClient = new SmtpClient(_host, _port))
                {
                    smtpClient.EnableSsl = _enableSsl;
                    smtpClient.UseDefaultCredentials = _defaultCredentials;
                    smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
                    if (_defaultCredentials == false)
                    {
                        smtpClient.Credentials = new NetworkCredential(_emailfrom, _password);
                    }
                    smtpClient.Send(_emailfrom, alıcılar, subject, content);
                }
                File.Delete(AppDomain.CurrentDomain.BaseDirectory + "mail.txt");
            }
            else { return; }
        }
        #endregion
    }
}
