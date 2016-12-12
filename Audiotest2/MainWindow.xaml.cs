using CSCore.CoreAudioAPI;
using CSCore.MediaFoundation;
using CSCore.SoundIn;
using CSCore.SoundOut;
using CSCore.Streams;
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

namespace Audiotest2
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            mmdevicesOut = MMDeviceEnumerator.EnumerateDevices(DataFlow.Render, DeviceState.Active).ToList();
            mmdevicesIn = MMDeviceEnumerator.EnumerateDevices(DataFlow.Capture, DeviceState.Active).ToList();
            comboBox.ItemsSource = mmdevicesOut.ToList().Concat(mmdevicesIn);
            comboBox_Copy.ItemsSource = mmdevicesOut.ToList().Concat(mmdevicesIn);
        }
        private WasapiCapture capture = null;
        private WasapiOut w = null;
        private List<MMDevice> mmdevicesOut = new List<MMDevice>();
        private List<MMDevice> mmdevicesIn = new List<MMDevice>();
 

        private void button_Click(object sender, RoutedEventArgs e) 
        {
            MMDevice dev = (MMDevice)comboBox.SelectedItem;
            if (mmdevicesOut.Contains(dev))
            {
                capture = new WasapiLoopbackCapture();
            }
            else
            {
                capture = new WasapiCapture();

            }
            capture.Device = dev;

            capture.Initialize();

            w = new WasapiOut();

            w.Device = (MMDevice)comboBox_Copy.SelectedItem;

            w.Initialize(new SoundInSource(capture) { FillWithZeros = true });
            
            capture.Start();
            w.Play();

        }

        private void button_Copy_Click(object sender, RoutedEventArgs e)
        {
            if (w != null && capture != null)
            {
                //stop recording
                w.Stop();
                capture.Stop();
                w.Dispose();
                w = null;
                capture.Dispose();
                capture = null;
            }
        }
    }
}
