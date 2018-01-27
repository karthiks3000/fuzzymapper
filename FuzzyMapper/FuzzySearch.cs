using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FuzzyMapper
{
    public static class FuzzySearch
    {             


        //---------------------------------------------------------------------
        /// <summary>
        /// Fuzzy searches a list of strings.
        /// </summary>
        /// <param name="word">
        /// The word to find.
        /// </param>
        /// <param name="wordList">
        /// A list of word to be searched.
        /// </param>
        /// <param name="fuzzyness">
        /// Ration of the fuzzyness. A value of 0.8 means that the 
        /// difference between the word to find and the found words
        /// is less than 20%.
        /// </param>
        /// <returns>
        /// The list with the found words.
        /// </returns>
        /// <example>
        /// 
        /// </example>
        public static List<string> Search(
            string word,
            List<string> wordList,
            double fuzzyness)
        {

            List<string> foundWords = new List<string>();

            foreach (string s in wordList)
            {
                // Calculate the Levenshtein-distance:
                int levenshteinDistance = LevenshteinDistanceExtensions.LevenshteinDistance(word, s);

                // Length of the longer string:
                int length = Math.Max(word.Length, s.Length);

                // Calculate the score:
                double score = 1.0 - (double)levenshteinDistance / length;

                // Match?
                if (score > fuzzyness)
                    foundWords.Add(s);
            }

            return foundWords;
        }
        //---------------------------------------------------------------------
        /// <summary>
        /// Fuzzy searches a list of strings using LINQ.
        /// </summary>
        /// <param name="word">
        /// The word to find.
        /// </param>
        /// <param name="wordList">
        /// A list of word to be searched.
        /// </param>
        /// <param name="fuzzyness">
        /// Ration of the fuzzyness. A value of 0.8 means that the 
        /// difference between the word to find and the found words
        /// is less than 20%.
        /// </param>
        /// <returns>
        /// The list with the found words.
        /// </returns>
        /// <example>
        /// 
        /// </example>
        public static List<string> Search_v2(string word, List<string> wordList, double fuzzyness)
        {

            List<string> foundWords =
                (
                    from s in wordList
                    let levenshteinDistance = LevenshteinDistanceExtensions.LevenshteinDistance(word, s)
                    let length = Math.Max(s.Length, word.Length)
                    let score = 1.0 - (double)levenshteinDistance / length
                    where score > fuzzyness
                    select s
                ).ToList();

            return foundWords;
        }

        //---------------------------------------------------------------------
        /// <summary>
        /// Fuzzy searches a dictionary<string,string> using LINQ.
        /// </summary>
        /// <param name="word">
        /// The word to find.
        /// </param>
        /// <param name="wordList">
        /// A dictionary of words to be searched.
        /// </param>
        /// <param name="fuzzyness">
        /// Ration of the fuzzyness. A value of 0.8 means that the 
        /// difference between the word to find and the found words
        /// is less than 20%.
        /// </param>
        /// <returns>
        /// The dictionary with the found words.
        /// </returns>
        /// <example>
        /// 
        /// </example>
        public static Dictionary<string, string> Search_v3(string word, Dictionary<string, string> wordList, double fuzzyness,string algorithm = "Levenshtein Distance")
        {

            Dictionary<string, string> foundWords;
            if (algorithm.Equals("Levenshtein Distance"))
            {
                foundWords =
                    (
                        from s in wordList
                        let levenshteinDistance = LevenshteinDistanceExtensions.LevenshteinDistance(word, s.Value)
                        let length = Math.Max(s.Value.Length, word.Length)
                        let score = 1.0 - (double)levenshteinDistance / length
                        where score > fuzzyness
                        select s
                    ).ToDictionary(t => t.Key, t => t.Value);
            }
            else if (algorithm.Equals("Dice Coefficient"))
            {
                foundWords =
                    (
                        from s in wordList
                        let score = DiceCoefficientExtensions.DiceCoefficient(word, s.Value)                        
                        where score > fuzzyness
                        select s
                    ).ToDictionary(t => t.Key, t => t.Value);
            }
            else if (algorithm.Equals("Longest Common Subsequence"))
            {
                foundWords =
                    (
                        from s in wordList
                        let score = LongestCommonSubsequenceExtensions.LongestCommonSubsequence(word, s.Value)
                        where score.Item2 > fuzzyness
                        select s
                    ).ToDictionary(t => t.Key, t => t.Value);
            }
            else if (algorithm.Equals("Double Metaphone"))
            {
                foundWords =
                    (
                        from s in wordList
                        let score = DoubleMetaphoneExtensions.DoubleMetaphoneCoefficient(word,s.Value)
                        where score > fuzzyness
                        select s
                    ).ToDictionary(t => t.Key, t => t.Value);
            }
            else
            {
                foundWords =
                    (
                        from s in wordList
                        where word.FuzzyEquals(s.Value, fuzzyness)
                        select s
                    ).ToDictionary(t => t.Key, t => t.Value);
            }

            return foundWords;
        }
    }
}
