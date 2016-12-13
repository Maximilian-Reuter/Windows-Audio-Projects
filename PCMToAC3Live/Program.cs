using CSCore.CoreAudioAPI;
using CSCore.SoundIn;
using CSCore.SoundOut;
using System;
using System.IO;
using Aften;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using CSCore.Streams;
using CSCore;
using CSCore.Codecs.WAV;

namespace PCMToAC3Live
{
    class Program
    {
        private static WasapiCapture capture = null;

        private static WasapiOut w = null;

        private static Stream fstream;
        private static byte[] fbuf;
        private static SingleBlockNotificationStream nStream;
        private static FrameEncoderFloat enc;
        private static byte[] trashBuf;
        private static IWaveSource final;
        private static WaveWriter _writer;
        private static long counter;
        private static Queue<float[]> sampleQueue = new Queue<float[]>();
        private static float[] queueBuf;

        static unsafe void Main(string[] args)

        {

            MMDevice dev = MMDeviceEnumerator.DefaultAudioEndpoint(DataFlow.Render, Role.Multimedia);

            capture = new WasapiLoopbackCapture();
            capture.Device = dev;
            capture.Initialize();

            SoundInSource soundInSource = new SoundInSource(capture);

            nStream = new SingleBlockNotificationStream(soundInSource.ToSampleSource());
            final = nStream.ToWaveSource();
            nStream.SingleBlockRead += NStream_SingleBlockRead;
            soundInSource.DataAvailable += encode;
            trashBuf = new byte[final.WaveFormat.BytesPerSecond / 2];

            Console.WriteLine($"sample rate:{capture.WaveFormat.SampleRate}");
            Console.WriteLine($"bits per sample:{capture.WaveFormat.BitsPerSample }");
            Console.WriteLine($"channels:{capture.WaveFormat.Channels }");
            Console.WriteLine($"bytes per sample:{capture.WaveFormat.BytesPerSample }");
            Console.WriteLine($"bytes per second:{capture.WaveFormat.BytesPerSecond }");
            Console.WriteLine($"AudioEncoding:{capture.WaveFormat.WaveFormatTag  }");


            EncodingContext context = FrameEncoder.GetDefaultsContext();
            context.Channels = 6;
            context.SampleRate = 44100;
            context.AudioCodingMode = AudioCodingMode.Front3Rear2;
            context.HasLfe = true;

            context.SampleFormat = A52SampleFormat.Float;
            enc = new FrameEncoderFloat(ref context);

            //_writer = new WaveWriter("test.ac3", final.WaveFormat);


            File.Create("test.ac3").Close();
            capture.Start();

            //w = new WasapiOut();
            //w.Device = (MMDevice)comboBox_Copy.SelectedItem;
            //w.Initialize(new SoundInSource(capture) { FillWithZeros = true });
            //w.Play();


            Task.Run(() => { encoderThread(); });
            //encodeSinus();

            Console.ReadLine();

            System.Environment.Exit(0);
        }

        private static void encoderThread()
        {
            fstream = File.Open("test.ac3", FileMode.Append);
            int i = 0;
            while (true)
            {
                if (sampleQueue.Count > 0)
                {
                    float[] samples = sampleQueue.Dequeue();
                    enc.Encode(samples, samples.Length / 6, (b, o, c) => fstream.Write(b, o, c));
                    if (i++ % 10 == 0) fstream.Flush();
                    //enc.Flush(fstream);
                }
            }
        }

        private static void NStream_SingleBlockRead(object sender, SingleBlockReadEventArgs e)
        {

            Console.WriteLine(String.Join(",", e.Samples.Select(x => Math.Round(x, 4).ToString())));


            if (queueBuf == null)
            {
                queueBuf = new float[44100 * 6];
            }


            queueBuf[counter * e.Samples.Length] = e.Samples[0];
            queueBuf[counter * e.Samples.Length + 1] = e.Samples[2];
            queueBuf[counter * e.Samples.Length + 2] = e.Samples[1];
            queueBuf[counter * e.Samples.Length + 3] = e.Samples[4];
            queueBuf[counter * e.Samples.Length + 4] = e.Samples[5];
            queueBuf[counter * e.Samples.Length + 5] = e.Samples[3];
            counter++;
            if (counter == 44100)
            {
                sampleQueue.Enqueue(queueBuf);
                queueBuf = null;
                counter = 0;

            }

            //fstream.Write( e.Samples.Select(x=>(byte) x).ToArray(), 0, 6);

            //fstream.Close();
            //if (e.Samples[0] != 0)
            //{
            //    Console.WriteLine("sample > 0");
            //}
        }

        private static unsafe void encode(object sender, DataAvailableEventArgs e)
        {
            int read;
            while ((read = final.Read(trashBuf, 0, trashBuf.Length)) > 0)
            {
                //_writer.Write(trashBuf, 0, read);
                //Console.WriteLine("read");
            }
        }
        private static unsafe void encodeSinus()
        {
            //float[] samplesSharp = new float[c->frame_size * c->channels];
            //double t = 0;
            //double tincr = 2 * Math.PI * 440 / c->sample_rate;
            //for (int i = 0; i < 200; i++)
            //{
            //    int got_output;
            //    fixed (AVPacket* pkt_fixed = &pkt)
            //    {
            //        ffmpeg.av_init_packet(pkt_fixed);
            //        pkt.data = null; // packet data will be allocated by the encoder
            //        pkt.size = 0;

            //        for (int j = 0; j < c->frame_size;j++)
            //        {
            //            float temp = (float)(Math.Sin(t) * 0.25);


            //            for (int k = 0; k < c->channels; k++)
            //                samplesSharp[c->frame_size * k + j] = temp;
            //            t += tincr;
            //        }

            //        Marshal.Copy(samplesSharp, 0, new IntPtr(samples), samplesSharp.Length);

            //        ffmpeg.avcodec_encode_audio2(c, pkt_fixed, frame, &got_output);
            //        if (got_output == 1)
            //        {
            //            byte* p = (byte*)pkt.data;
            //            for (int j = 0; j < pkt.size; j++)
            //            {

            //                fbuf[j] = *p;

            //                p++;
            //            }
            //            fstream.Write(fbuf, 0, pkt.size);
            //            ffmpeg.av_packet_unref(pkt_fixed);
            //        }
            //    }
            //}
        }

    }
}
