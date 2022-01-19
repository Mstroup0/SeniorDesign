﻿using System;
using System.Linq;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace WordPredictionLibrary.Core
{
	public class NextWordFrequencyDictionary
	{
		public int DistinctWordCount { get { return _internalDictionary.Count; } }
		public decimal TotalWordCount { get { return _internalDictionary.Values.Sum(); } }

		internal Dictionary<Word, decimal> _internalDictionary = null;
		private static decimal noMatchValue = 0;

		#region Constructors

		public NextWordFrequencyDictionary()
		{
			_internalDictionary = new Dictionary<Word, decimal>();
		}

		public NextWordFrequencyDictionary(Dictionary<Word, decimal> dictionary)
			: this()
		{
			_internalDictionary = dictionary;
		}

		#endregion

		#region ToString Override

		public override string ToString()
		{
			List<Tuple<Word, decimal>> tupleList = OrderByFrequencyDescending()
														.Select(kvp => new Tuple<Word, decimal>(kvp.Key, kvp.Value))
														.ToList();
			Tuple<Word, decimal> first = tupleList.FirstOrDefault();

			if (first == null) { return "first == null"; }

			int padding = 0;
			decimal counter = first.Item2;

			StringBuilder result = new StringBuilder();
			while (counter-- > 0)
			{
				List<Tuple<Word, decimal>> words = tupleList.Where(t => t.Item2 == counter)
															.Select(t => new Tuple<Word, decimal>(t.Item1, t.Item2))
															.ToList();
				if (words == null || words.Count < 1)
				{
					continue;
				}

				padding++;

				result.AppendFormat(new string(Enumerable.Repeat<char>(' ', padding).ToArray()));

				foreach(Tuple<Word, decimal> tuple in words)
				{
					result.AppendFormat("{0}:{1}   ", tuple.Item2, tuple.Item1.Value);
				}

				result.AppendLine("");
			}
			return result.ToString();
		}

		#endregion

		#region Dictionary Altering Methods

		public void Add(Word word)
		{
			if (this.Contains(word))
			{
				_internalDictionary[word] += 1;
			}
			else
			{
				_internalDictionary.Add(word, 1);
			}
		}

		public decimal this[Word key]
		{
			get { return this.Contains(key) ? _internalDictionary[key] : noMatchValue; }
		}

		public bool Contains(Word key)
		{
			return _internalDictionary.ContainsKey(key);
		}

		public bool Contains(string key)
		{
			string lowerKey = key.TryToLower();
			return _internalDictionary.Any(kvp => kvp.Key.Equals(lowerKey));
		}

		#endregion

		#region Suggest Methods

		public string GetNextWord()
		{
			return GetNextWordByFrequencyDescending().FirstOrDefault();
		}

		public IEnumerable<string> TakeTop(int count)
		{
			if (count == -1)  
			{
				return GetNextWordByFrequencyDescending();
			}
			return GetNextWordByFrequencyDescending().Take(count);
		}

		public IEnumerable<string> Next4(int count)
        {

			return GetNextWordByFrequencyDescending().Take(count);
		}



		public IEnumerable<string> GetNextWordByFrequencyDescending()
		{
			return OrderByFrequencyDescending().Select(kvp => kvp.Key.Value);    /// Edit here to add more words. 
		}

		private IOrderedEnumerable<KeyValuePair<Word, decimal>> _orderedDictionary = null;
		public IOrderedEnumerable<KeyValuePair<Word, decimal>> OrderByFrequencyDescending()
		{
			// If we haven't set FrequencyDictionary yet OR it is out of date (dict has more entries)
			if (_orderedDictionary == null || _internalDictionary.Count > _orderedDictionary.Count())
			{
				_orderedDictionary = _internalDictionary.OrderByDescending(kvp => kvp.Value);
			}

			return _orderedDictionary;
		}

		#endregion


		/// <summary>
		/// The frequency is a value between 0 and 1 that is a proportional fraction of 
		/// the number of times it was observed out of all the data observed.
		/// </summary>
		public Dictionary<Word, decimal> GetFrequencyDictionary()
		{
			decimal baseProbability = 1 / TotalWordCount;

			Dictionary<Word, decimal> frequencyDict = _internalDictionary.ToDictionary(k => k.Key, v => baseProbability * v.Value);
			IOrderedEnumerable<KeyValuePair<Word, decimal>> orderedFreqKvp = frequencyDict.OrderByDescending(kvp => kvp.Value);

			return orderedFreqKvp.ToDictionary(k => k.Key, v => v.Value);
		}
	}
}
