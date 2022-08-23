using System;
using System.Collections.Generic;
using RestSharp;

namespace UrlopyDelegacje
{
	public class Urlop
	{
		private Form1 form1 = new Form1();

		public long ID { get; set; }

		public DateTime Od { get; set; }

		public DateTime Do { get; set; }

		public int DniIlosc { get; set; }

		public string Comments { get; set; }

		public string WniosekPath { get; set; }
		public string Swieto { get; set; }
		public bool Delegacja { get; set; }
		public double Zwrot { get; set; }

		public Urlop(DateTime aOD, DateTime aDO, bool aDelegacja)
		{
			ID = Math.Abs(DateTime.Now.ToBinary());
			Od = aOD;
			Do = aDO;
			if(aDelegacja)
            {
				DniIlosc = (Do - Od).Days + 1;
			}
            else
            {
				DniIlosc = (Do - Od).Days + 1 - DniWeekend(aOD, aDO);
			}
			if (DniIlosc <= 0 || aOD.Year != aDO.Year)
			{
				throw new IndexOutOfRangeException();
			}
			Comments = "";
			WniosekPath = "";
			Delegacja = aDelegacja;
			Zwrot = 0;
			//Swieto = "";
		}

		public Urlop()
		{
		}

		public int DniWeekend(DateTime from, DateTime thru)
		{
			int num = 0;
			DateTime item = from.Date;
			while (item.Date <= thru.Date)
			{
				if (item.DayOfWeek == DayOfWeek.Saturday || item.DayOfWeek == DayOfWeek.Sunday)
				{
					num++;
				}else if(form1.CheckSwieto(item.Year, item.Month, item.Day).Content.Contains("National"))
				{
					num++;
					Swieto += item.Date.Day.ToString() + "." + item.Date.Month.ToString() + ",";
				}
				item = item.AddDays(1.0);
			}
			return num;
		}

		public List<DateTime> GetSwieta(int year)
		{		
			List<DateTime> list = new List<DateTime>();
			try
			{
				list.Add(new DateTime(year, 1, 1));
				list.Add(new DateTime(year, 4, 10));
				list.Add(new DateTime(year, 4, 13));
				list.Add(new DateTime(year, 5, 1));
				list.Add(new DateTime(year, 6, 1));
				list.Add(new DateTime(year, 10, 3));
				list.Add(new DateTime(year, 12, 25));
				list.Add(new DateTime(year, 12, 26));
			}
			catch
			{
				list = form1.FillSwieta();
			}
			return list;

		}
	}
}
