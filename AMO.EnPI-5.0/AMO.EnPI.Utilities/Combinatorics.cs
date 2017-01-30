using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AMO.EnPI.AddIn.Utilities
{
    public class Combination
    {
        // AMO.EnPI.Utilities.Combination
        //
        // This class contains the combinatorics processing components for the EnPI Addin. 
        //
        // Combination(n,k) is all possible k-sized combinations of n values
        // Choose(n,k) is a count of all possible k-sized combinations of n values
        // Combination(n,k).Element(m) is the mth element of all possible k-sized combinations of n values

        private int n = 0;
        private int k = 0;
        private int[] data = null;

        public Combination(int n, int k)
        {
            if (n < 0 || k < 0) // normally n >= k
                throw new Exception("Negative parameter in constructor");

            this.n = n;
            this.k = k;
            this.data = new int[k];
            for (int i = 0; i < k; ++i)
                this.data[i] = i;
        } // Combination(n,k)

        public Combination(int n, int k, int[] a) // Combination from a[]
        {
            if (k != a.Length)
                throw new Exception("Array length does not equal k");

            this.n = n;
            this.k = k;
            this.data = new int[k];
            for (int i = 0; i < a.Length; ++i)
                this.data[i] = a[i];

            if (!this.IsValid())
                throw new Exception("Bad value from array");
        } // Combination(n,k,a)

        public bool IsValid()
        {
            if (this.data.Length != this.k)
                return false; // corrupted

            for (int i = 0; i < this.k; ++i)
            {
                if (this.data[i] < 0 || this.data[i] > this.n - 1)
                    return false; // value out of range

                for (int j = i + 1; j < this.k; ++j)
                    if (this.data[i] >= this.data[j])
                        return false; // duplicate or not lexicographic
            }

            return true;
        } // IsValid()

        public override string ToString()
        {
            string s = "{ ";
            for (int i = 0; i < this.k; ++i)
                s += this.data[i].ToString() + " ";
            s += "}";
            return s;
        } // ToString()

        public int[] ToArray()
        {
            int[] a = new int[this.k];
            int val;
            for (int i = 0; i < this.k; ++i)
            {
             if (int.TryParse(this.data[i].ToString(), out val) )
                 a[i] = val;
            }
            return a;
        }

        public Combination Successor()
        {
            if (this.data[0] == this.n - this.k)
                return null;

            Combination ans = new Combination(this.n, this.k);

            int i;
            for (i = 0; i < this.k; ++i)
                ans.data[i] = this.data[i];

            for (i = this.k - 1; i > 0 && ans.data[i] == this.n - this.k + i; --i)
                ;

            ++ans.data[i];

            for (int j = i; j < this.k - 1; ++j)
                ans.data[j + 1] = ans.data[j] + 1;

            return ans;
        } // Successor()

        // Implements "n Choose k" -- the number of possible combinations of n that are k long
        // nCk = n! / (k!(n-k)!)
        public static int Choose(int n, int k)
        {
            if (n < 0 || k < 0)
                throw new Exception("Invalid negative parameter in Choose()");
            if (n < k)
                return 0;  // special case
            if (n == k)
                return 1;

            int delta, iMax;

            if (k < n - k) // ex: Choose(100,3)
            {
                delta = n - k;
                iMax = k;
            }
            else         // ex: Choose(100,97)
            {
                delta = k;
                iMax = n - k;
            }

            int ans = delta + 1;

            for (int i = 2; i <= iMax; ++i)
            {
                checked { ans = (ans * (delta + i)) / i; }
            }

            return ans;
        } // Choose()
        
        // return the mth lexicographic element of combination C(n,k)
        public Combination Element(int m)
        {
            int[] ans = new int[this.k];

            int a = this.n;
            int b = this.k;
            int x = (Choose(this.n, this.k) - 1) - m; // x is the "dual" of m

            for (long i = 0; i < this.k; ++i)
            {
                ans[i] = LargestV(a, b, x); // largest value v, where v < a and vCb < x    
                x = x - Choose(ans[i], b);
                a = ans[i];
                b = b - 1;
            }

            for (long i = 0; i < this.k; ++i)
            {
                ans[i] = (n - 1) - ans[i];
            }

            return new Combination(this.n, this.k, ans);
        } // Element()


        // return largest value v where v < a and  Choose(v,b) <= x
        private static int LargestV(int a, int b, int x)
        {
            int v = a - 1;

            while (Choose(v, b) > x)
                --v;

            return v;
        } // LargestV()


    } // Combination class
}
