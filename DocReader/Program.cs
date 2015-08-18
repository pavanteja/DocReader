using OpenMcdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


//OpenMDCF from nuget
namespace DocReader
{
    class Program
    {
        static void Main(string[] args)
        {
            string text = "";
            using (CompoundFile wordDoc = new CompoundFile("E:\\sample.doc"))
            {
                CFStream cf = wordDoc.RootStorage.GetStream("WordDocument");

                //cf.Seek(512, SeekOrigin.Begin);//Trim 512-Byte header of Compound File Binary Format
                byte[] fib = cf.GetData();// new byte[1472];
                                          //cf.Read(fib, 0, 1472);

                ByteExtension.print("wIdent", fib[0], fib[1]);
                ByteExtension.print("nFib", fib[2], fib[3]);
                ByteExtension.print("unused (2 bytes)", fib[4], fib[5]);
                ByteExtension.print("lid", fib[6], fib[7]);
                ByteExtension.print("pnNext ", fib[8], fib[9]);
                Console.WriteLine("A - fDot (1 bit) -" + fib[10].bit(0));
                Console.WriteLine("B - fGlsy  (1 bit) -" + fib[10].bit(1));
                Console.WriteLine("C - fComplex  (1 bit) -" + fib[10].bit(2));
                Console.WriteLine("D - fHasPic  (1 bit) -" + fib[10].bit(3));
                Console.WriteLine("E - cQuickSaves  (4 bit) -" + fib[10].bit(4) + fib[10].bit(5) + fib[10].bit(6) + fib[10].bit(7));
                Console.WriteLine("F - fEncrypted  (1 bit) -" + fib[11].bit(0));
                Console.WriteLine("G - fWhichTblStm  (1 bit) -" + fib[11].bit(1));
                Console.WriteLine("H - fReadOnlyRecommended  (1 bit) -" + fib[11].bit(2));
                Console.WriteLine("I - fWriteReservation  (1 bit) -" + fib[11].bit(3));
                Console.WriteLine("J - fExtChar  (1 bit) -" + fib[11].bit(4));
                Console.WriteLine("K - fLoadOverride  (1 bit) -" + fib[11].bit(5));
                Console.WriteLine("L - fFarEast  (1 bit) -" + fib[11].bit(6));
                Console.WriteLine("M - fObfuscated  (1 bit) -" + fib[11].bit(7));
                ByteExtension.print("nFibBack ", fib[12], fib[13]);
                ByteExtension.print("lKey (2 of 4 Bytes) ", fib[14], fib[15]);
                ByteExtension.print("lKey (2 of 4 Bytes) ", fib[16], fib[17]);
                Console.WriteLine("envr (1 byte) 0x" + fib[18]);
                Console.WriteLine("N - fMac (1 bit) -" + fib[19].bit(0));
                Console.WriteLine("O - fEmptySpecial (1 bit) -" + fib[19].bit(1));
                Console.WriteLine("P - fLoadOverridePage (1 bit) -" + fib[19].bit(2));
                Console.WriteLine("Q - reserved1 (1 bit) -" + fib[19].bit(3));
                Console.WriteLine("R - reserved2 (1 bit) -" + fib[19].bit(4));
                Console.WriteLine("S - fSpare0 (3 bits) -" + fib[19].bit(5) + fib[19].bit(6) + fib[19].bit(7));
                ByteExtension.print("reserved3 (2 bytes) -", fib[20], fib[21]);
                ByteExtension.print("reserved4 (2 bytes) -", fib[22], fib[23]);
                ByteExtension.print("reserved5 (2 of 4 bytes) -", fib[24], fib[25]);
                ByteExtension.print("reserved5 (2 of 4 bytes) -", fib[26], fib[27]);
                ByteExtension.print("reserved6 (2 of 4 bytes) -", fib[28], fib[29]);
                ByteExtension.print("reserved6 (2 of 4 bytes) -", fib[30], fib[31]);




                UInt32 fcClx = System.BitConverter.ToUInt32(fib, 0x01A2);
                UInt32 lcbClx = System.BitConverter.ToUInt32(fib, 0x01A6);

                bool flag1Table = ((fib[0x000A] & 0x0200) == 0x0200);

                string tableStreamName = "1Table";
                if (flag1Table) tableStreamName = "1Table";


                CFStream tableStream = wordDoc.RootStorage.GetStream(tableStreamName);
                int count = 0;
                byte[] clx = tableStream.GetData().Skip((int)fcClx).Take((int)lcbClx).ToArray();


                byte[] bytes = new byte[20];
                int pos = 0;
                bool goOn = true;
                while (goOn)
                {
                    byte typeEntry = clx[pos];
                    if (typeEntry == 2)
                    {
                        goOn = false;
                        Int32 lcbPieceTable = System.BitConverter.ToInt32(clx, pos + 1);
                        byte[] pieceTable = new byte[lcbPieceTable];
                        Array.Copy(clx, pos + 5, pieceTable, 0, pieceTable.Length);
                        int pieceCount = (lcbPieceTable - 4) / 12;

                        for (int i = 0; i < pieceCount; i++)
                        {
                            Int32 cpStart = System.BitConverter.ToInt32(pieceTable, i * 4);
                            Int32 cpEnd = System.BitConverter.ToInt32(pieceTable, (i + 1) * 4);

                            byte[] pieceDescriptor = new byte[8];
                            int offsetPieceDescriptor = ((pieceCount + 1) * 4) + (i * 8);
                            Array.Copy(pieceTable, offsetPieceDescriptor, pieceDescriptor, 0, 8);

                            UInt32 fcValue = System.BitConverter.ToUInt32(pieceDescriptor, 2);
                            bool isANSI = ((fcValue & 0x40000000) == 0x40000000);
                            uint fc = (fcValue & 0xBFFFFFFF);

                            Encoding encoding = Encoding.GetEncoding(1252);
                            Int32 cb = cpEnd - cpStart;
                            if (!isANSI)
                            {
                                encoding = Encoding.Unicode;
                                cb *= 2;
                            }

                            byte[] bytesOfText = new byte[cb];
                            var take = (int)fc;
                            bytesOfText = cf.GetData().Skip(take/2).Take(cb).ToArray();

                            //wordDocumentStream.Read(bytesOfText, bytesOfText.Length, fc);
                            text += encoding.GetString(bytesOfText);
                            System.IO.StreamWriter file = new System.IO.StreamWriter("D:\\test.txt");
                            file.WriteLine(text);
                            file.Close();
                        }
                    }
                    else if (typeEntry == 1)
                    {
                        pos = pos + 1 + 1 + clx[pos + 1];
                    }
                    else
                    {
                        goOn = false;
                    }
                }

                
                Console.WriteLine(text);
                Console.ReadKey();

            }
        }


    }

    public static class ByteExtension
    {
        public static void print(this String name, byte b1, byte b2)
        {
            Console.WriteLine(String.Format("{0} - 0x{2:X2}{1:X2}", name, b1, b2));
        }
        public static string bit(this byte b1, int bitnumber)
        {
            if ((b1 & (1 << bitnumber - 1)) != 0)
            {
                return "1";
            }
            else
            {
                return "0";
            }
        }
    }
}
