using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    public static class StringDistance
    {
        /// <summary>
        /// Damerau-Levenshtein Distance Algorithm
        /// 
        /// The function determines the difference between two strings.
        /// The distance is the minimum number of edits (insertions,
        /// deletions, substitutions, and transpositions of two characters)
        /// needed to change one string into the other.
        /// </summary>
        /// <param name="s"></param>
        /// <param name="t"></param>
        /// <returns></returns>
        public static int GetStringDistance(string s, string t)
        {
            int mtrxHeight = s.Length + 1;
            int mtrxWidth = t.Length + 1;
            int[,] matrix = new int[mtrxHeight, mtrxWidth]; // main matrix for algorithm

            for (int i = 0; i < mtrxHeight; i++)
                matrix[i, 0] = i;
            for (int j = 0; j < mtrxWidth; j++)
                matrix[0, j] = j;

            for (int height = 1; height < mtrxHeight; height++)
            {
                for (int width = 1; width < mtrxWidth; width++)
                {
                    int cost = (s[height - 1] == t[width - 1]) ? 0 : 1; // determine cost

                    int distance = Math.Min(matrix[height - 1, width] + 1, // deletion
                        Math.Min(matrix[height, width - 1] + 1, // insertion
                        matrix[height - 1, width - 1] + cost)); // substitution

                    if (height > 1 && width > 1 && s[height - 1] == t[width - 2] && s[height - 2] == t[width - 1])
                    {
                        distance = Math.Min(distance,
                            matrix[height - 2, width - 2] + cost); // transposition
                    }

                    matrix[height, width] = distance;
                }
            }

            return matrix[mtrxHeight - 1, mtrxWidth - 1];
        }
    }
}
