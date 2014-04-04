using System;
using System.Collections.Generic;
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
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Management;
using System.IO;
using System.Security;
using System.ComponentModel;
using Microsoft.Win32;

namespace WinInventory
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<String, Dictionary<String, String>> productKeyCollection;
        private List<string> devices = new List<string>();
        private readonly BackgroundWorker bw = new BackgroundWorker();
        private Dictionary<string, string> detailedInformation = new Dictionary<string,string>();

        public MainWindow()
        {
            bw.DoWork += this.background_work;
            bw.RunWorkerCompleted += this.background_work_completed;
            InitializeComponent();
        }

        private void getKeys(object sender, RoutedEventArgs e)
        {
            this.result.ItemsSource = null;
            this.result.Items.Refresh();
            this.devicename.SelectionChanged -= displayDetails;

            this.productKeyCollection = new Dictionary<string, Dictionary<string, string>>();
            IEnumerable<CheckBox> c = picker.Children.OfType<CheckBox>();
            List<string> choices = new List<string>();
            foreach (CheckBox cb in c)
            {
                if ((bool)cb.IsChecked)
                {
                    choices.Add(cb.Name);
                }
            }
            
            this.Cursor = Cursors.Wait;
            bw.RunWorkerAsync(choices);
        }

        private void displayDetails(object sender, SelectionChangedEventArgs e)
        {
            this.result.ItemsSource = productKeyCollection[this.devicename.SelectedItem.ToString()];
            this.result.Columns[0].Header = "Software Package";
            this.result.Columns[1].Header = "Product Key";
            this.result.Items.Refresh();
        }
    
        private static Dictionary<String, String> getProductKeys(string product, int version, string computername)
        {
            Dictionary<String, String> productkeys = new Dictionary<string,string>();
            string productname;
            byte[] digitalProductId;
            RegistryKey key;
            key = RegistryKey.OpenRemoteBaseKey(RegistryHive.LocalMachine, computername, RegistryView.Registry64);
            try
            {
                switch (product)
                {
                    case "win":
                        key = key.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false);
                        productname = (string)key.GetValue("ProductName");
                        digitalProductId = (byte[])key.GetValue("DigitalProductId");
                        productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                        break;
                    case "mso":
                        switch (version)
                        {
                            case 11:
                                key = key.OpenSubKey(@"SOFTWARE\Microsoft\Office\11.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                key = key.OpenSubKey(@"SOFTWARE\wow6432node\Microsoft\Office\11.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                break;
                            case 12:
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                key = key.OpenSubKey(@"SOFTWARE\wow6432node\Microsoft\Office\12.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            System.Diagnostics.Trace.WriteLine(key.GetSubKeyNames()[i]);
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                break;
                            case 14:
                                key = key.OpenSubKey(@"SOFTWARE\Microsoft\Office\14.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");
                                        
                                        if(productname==null | digitalProductId==null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }
                                        
                                    }
                                }
                                key = key.OpenSubKey(@"SOFTWARE\wow6432node\Microsoft\Office\14.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductId");
                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                    }
                                }
                                break;
                            case 15:
                                key = key.OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                key = key.OpenSubKey(@"SOFTWARE\wow6432node\Microsoft\Office\15.0\Registration", false);
                                if (key != null)
                                {
                                    for (int i = 0; i < key.GetSubKeyNames().Length; i++)
                                    {
                                        productname = (string)key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("ConvertToEdition");
                                        digitalProductId = (byte[])key.OpenSubKey(key.GetSubKeyNames()[i], false).GetValue("DigitalProductID");

                                        if (productname == null | digitalProductId == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            productkeys.Add(productname, decodeMicrosoftProductID(digitalProductId));
                                        }

                                    }
                                }
                                break;
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e.Message);
            }
            return productkeys;
        }

        private static string decodeMicrosoftProductID(byte[] digitalProductId) {
            if(digitalProductId == null)
            {
                return "Product key not available.";
            }

            String productkey;
            const int KeyOffset = 52;
            string Chars = "BCDFGHJKMPQRTVWXY2346789";
            try
            {
                byte isWin8 = (byte)((digitalProductId[66] / 6) & 1);
                digitalProductId[66] = (byte)((digitalProductId[66] & 0xF7) | ((isWin8 & 2) * 4));
                int iteration = 24;
                string keyoutput = "";
                int last = 0;
                do
                { 
                    int current = 0;
                    int X = 14;
                    do
                    {
                        current = current * 256;
                        current = digitalProductId[X + KeyOffset] + current;
                        digitalProductId[X + KeyOffset] = (byte)(current / 24);
                        current = current%24;
                        --X;
                    }
                    while(X>=0);
                    iteration--;
                    keyoutput = Chars.Substring(current,1) + keyoutput;             
                    last = current;
                }
                while(iteration>=0);
                if (isWin8 == 1)
                {
                    string keypart1 = keyoutput.Substring(1, last);
                    keyoutput = keyoutput.Substring(1);
                    int pos = keyoutput.IndexOf(keypart1, StringComparison.OrdinalIgnoreCase);
                    if (pos > -1)
                    {
                        keyoutput = keyoutput.Substring(0, pos) + keypart1 + "N" + keyoutput.Substring(pos+keypart1.Length);
                    }
                    if (last == 0) 
                    {
                        keyoutput = "N" + keyoutput;
                    }
                }
                String a = keyoutput.Substring(0, 5);
                String b = keyoutput.Substring(5, 5);
                String c = keyoutput.Substring(10, 5);
                String d = keyoutput.Substring(15, 5);
                String e = keyoutput.Substring(20, 5);
                
                productkey = a + "-" + b + "-" + c + "-" + d + "-" + e;
            }
            catch (IOException e)
            {
                productkey = e.Message + "\nHint: Remote Registry Service has to be enabled.";
            }
            catch (SecurityException e)
            {
                productkey = e.Message;
            }
            catch (UnauthorizedAccessException e)
            {
                productkey = e.Message;
            }
            catch (Exception e)
            {
                productkey = "Unspecified error occured.\n" + e.Message;
            }
            return productkey;
        }

        private void background_work(object sender, DoWorkEventArgs e)
        {
            List<string> choices = (List<string>)e.Argument;
            var osSelected = choices.Any(m => m.Contains("win"));

            List<string> targetDevices = new List<string>();
            DirectorySearcher ds = new DirectorySearcher();
            ds.Filter = ("(objectClass=computer)");

            foreach (SearchResult item in ds.FindAll())
            {
                String osVersion = item.GetDirectoryEntry().Properties["operatingSystemVersion"].Value.ToString();

                if (osSelected) {
                    Parallel.ForEach(choices, choice =>
                    {
                        if (choice.Equals("win63"))
                        {
                            if (osVersion.Equals("6.3 (9600)"))
                            {
                                targetDevices.Add(item.GetDirectoryEntry().Properties["dNSHostName"].Value.ToString());
                            }

                        }
                        else if (choice.Equals("win62"))
                        {
                            if (osVersion.Equals("6.2 (9200)"))
                            {
                                targetDevices.Add(item.GetDirectoryEntry().Properties["dNSHostName"].Value.ToString());
                            }

                        }
                        else if (choice.Equals("win61"))
                        {
                            if (osVersion.Equals("6.1 (7601)"))
                            {
                                targetDevices.Add(item.GetDirectoryEntry().Properties["dNSHostName"].Value.ToString());
                            }

                        }
                        else if (choice.Equals("win60"))
                        {
                            if (osVersion.Equals("6.0 (6000)"))
                            {
                                targetDevices.Add(item.GetDirectoryEntry().Properties["dNSHostName"].Value.ToString());
                            }

                        }
                    });
                }
                else
                {
                    targetDevices.Add(item.GetDirectoryEntry().Properties["dNSHostName"].Value.ToString());
                }
            }
            
            List<string> notReachableDevices = new List<string>();
            
            if (!osSelected)
            {
                Parallel.ForEach(targetDevices, computername =>
                {
                    Dictionary<string, string> productKeys = new Dictionary<string, string>();
                    productKeyCollection.Add(computername, productKeys);
                });
            }
            else
            {
                Parallel.ForEach(targetDevices, computername =>
                {
                    Dictionary<string, string> productKeys = new Dictionary<string, string>();
                    try
                    {
                        Dictionary<string, string> winkey = getProductKeys("win", 0, computername);
                        var result = productKeys.Concat(winkey).ToDictionary(el => el.Key, el => el.Value);
                        productKeyCollection.Add(computername, result);
                    }
                    catch (IOException)
                    {
                        notReachableDevices.Add(computername);
                    }
                });
            }
            targetDevices.RemoveAll(device => notReachableDevices.Contains(device));
            
            Parallel.ForEach(choices, choice =>
            {
                if (choice.Contains("mso"))
                {
                    Parallel.ForEach(targetDevices, computername =>
                    {
                        try
                        {
                            Dictionary<string, string> productKey = getProductKeys("mso", Int32.Parse(choice.Substring(3)), computername);
                            
                            ManagementScope scope = new ManagementScope("\\\\" + computername + "\\root\\cimv2");
                            scope.Connect();
                            ObjectQuery q = new ObjectQuery("SELECT Name, PartialProductKey FROM SoftwareLicensingProduct WHERE Name Like 'Office%' AND PartialProductKey <> null");
                            ManagementObjectSearcher s = new ManagementObjectSearcher(scope, q);

                            foreach (ManagementObject mo in s.Get())
                            {
                               productKey.Add(mo.Properties["Name"].Value.ToString(), "XXXXX-XXXXX-XXXXX-XXXXX-"+mo.Properties["PartialProductKey"].Value.ToString());
                            }
                            var result = productKeyCollection[computername].Concat(productKey).ToDictionary(el => el.Key, el => el.Value);
                            productKeyCollection[computername] = result;  
                        }
                        catch
                        {

                        }
                    });
                }
            });
            e.Result = productKeyCollection;
        }

        private void background_work_completed(object sender, RunWorkerCompletedEventArgs e)
        {
            Dictionary<string, Dictionary<string, string>> productKeyCollection = (Dictionary<string, Dictionary<string, string>>)e.Result;
            Dictionary<string, int> totalCount = new Dictionary<string, int>();
            this.productKeyCollection = productKeyCollection;
            this.devicename.ItemsSource = productKeyCollection.Keys;
            this.devicename.SelectedIndex = -1;
            this.devicename.Visibility = System.Windows.Visibility.Visible;
            this.devicename.Items.Refresh();
            this.devicename.SelectionChanged += displayDetails;

            totalCount.Add("Windows Vista", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("Windows Vista")).Count());
            totalCount.Add("Windows 7", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("Windows 7")).Count());
            totalCount.Add("Windows 8", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Substring(0, 10).Equals("Windows 8 ")).Count());
            totalCount.Add("Windows 8.1", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Substring(0, 11).Equals("Windows 8.1")).Count());
            totalCount.Add("MS Office 2003", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("2003")).Count());
            totalCount.Add("MS Office 2007", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("2007")).Count());
            totalCount.Add("MS Office 2010", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("2010")).Count());
            totalCount.Add("MS Office 2013", productKeyCollection.Values.SelectMany(v => v).Where(v => v.Key.Contains("2013")).Count());
            this.result.ItemsSource = totalCount;

            this.Cursor = Cursors.Arrow;
        }
    }
}
