using System.Text;

namespace GeneralReport
{
    class PasswordTool
    {
        private int[] mKey;
        private int mSeed;

        public PasswordTool(string key)
        {
            char[] c = key.ToCharArray();
            mSeed = 0;
            for (int i = 0; i < key.Length; i++)
            {
                mSeed += (int)c[i] * (i + 1) % 177;
            }
            mKey = new int[30];
            for (int i = 0; i < 30; i++)
            {
                mKey[i] = mSeed * (i + 1) % (127 - i);
            }
        }

        public string Encrypt(string str)
        {
            char[] ch = str.ToCharArray();
            int code;
            int cnt = 0;
            StringBuilder AdjPos = new StringBuilder();
            StringBuilder tmp = new StringBuilder();

            for (int i = 0; i < ch.Length; i++)
            {
                code = (int)ch[i] ^ mKey[i % 30];
                if (code < 35)
                {
                    code += 35;
                    cnt++;
                    AdjPos.Append((char)(i + 36));
                }
                tmp.Append((char)code);
            }

            StringBuilder sb = new StringBuilder();
            sb.Append((char)(cnt + 35));
            sb.Append(AdjPos);
            sb.Append(tmp);
            return sb.ToString();
        }

        public string Decrypt(string src)
        {
            char[] ch = src.ToCharArray();
            int cnt = (int)ch[0] - 35;
            char[] AdjPos = src.Substring(1, cnt).ToCharArray();
            char[] data = src.Substring(cnt + 1).ToCharArray();
            int pos;
            for (int i = 0; i < cnt; i++)
            {
                pos = (int)AdjPos[i] - 35;
                data[pos - 1] = (char)((int)data[pos - 1] - 35);
            }
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sb.Append((char)((int)data[i] ^ mKey[i % 30]));
            }
            return sb.ToString();
        }
    }
}
