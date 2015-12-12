using System;
using System.Linq;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;
using System.Collections.Generic;

namespace ConsoleTester
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                cout("Started...\n");

                //Fun1();
                Fun2();

            }
            catch (Exception exc)
            {
                cout("ERROR", exc.ToString());
            }

            if (args.Length <= 0)
            {
                cout("\n\nDone. Press any key.");
                Console.ReadLine();
            }
            else
            {
                cout("\n\nDone.");
            }
        }

        /// <summary>
        /// </summary>
        private static void Fun2()
        {
            var siteUrl = "http://sp.bandr.com/sites/Zen20";

            using (SPSite site = new SPSite(siteUrl))
            {
                var txsn = new TaxonomySession(site, true);

                //cout(txsn.TermStores.Count);

                var tstore = txsn.TermStores.First();

                //cout(tstore.Name);

                var tsetcol = tstore.GetTermSets("Books", 1033);
                var tset = tsetcol[0];

                var parentEntityTerm = tset.Terms["Children"];

                var lstSkipTerms = new List<Guid>();
                if (!string.IsNullOrEmpty(parentEntityTerm.CustomSortOrder))
                {
                    var guids = parentEntityTerm.CustomSortOrder.Split(":".ToCharArray());

                    foreach (var guid in guids)
                    {
                        lstSkipTerms.Add(new Guid(guid));
                    }
                }

                foreach (var guid in lstSkipTerms)
                {
                    var term = parentEntityTerm.TermSet.GetTerm(guid);
                    cout(term.Name, term.Id);
                }

                foreach (var term in parentEntityTerm.Terms)
                {
                    if (!lstSkipTerms.Contains(term.Id))
                    {
                        cout(term.Name, term.Id);
                    }
                }


            }
        }

        /// <summary>
        /// </summary>
        private static void Fun1()
        {
            var siteUrl = "http://sp.bandr.com";

            using (SPSite site = new SPSite(siteUrl))
            {
                var txsn = new TaxonomySession(site, true);

                //cout(txsn.TermStores.Count);

                var tstore = txsn.TermStores.First();

                //cout(tstore.Name);

                var tset = tstore.GetTermSets("Books", 1033);

                var terms = tstore.GetTerms(TermSetItem.NormalizeName("Garfield & \"Friends\""), false);
                cout(terms.Count);

                terms = tstore.GetTerms("Garfield & \"Friends\"", false);
                cout(terms.Count);

                cout();
                
                // both are found, with or without the call to the normalizename

                var term = tstore.GetTerm(new Guid("d06ce881-a36e-4d95-952b-8cd92e19fc95"));

                cout("term", term.GetDefaultLabel(1033));

                var label = term.GetDefaultLabel(1033);

                cout("original", label);
                label = label.Replace('&', '1'); // no impact
                label = label.Replace('"', '2'); // no impact

                label = label.Replace(Convert.ToChar(char.ConvertFromUtf32(65286)), '+');
                label = label.Replace(Convert.ToChar(char.ConvertFromUtf32(65282)), '\'');
                cout("fixed", label);

                cout("from function", MmdDenormalize(term.GetDefaultLabel(1033)));

                cout("as char codes:");
                foreach(char c in MmdDenormalize(term.GetDefaultLabel(1033)).ToCharArray())
                {
                    cout(" - ", c, (int) c);
                }


            }
        }

        /// <summary>
        /// </summary>
        static string MmdDenormalize(object o)
        {
            return GenUtil.SafeTrim(o)
                .Replace(Convert.ToChar(char.ConvertFromUtf32(65286)), '&')
                .Replace(Convert.ToChar(char.ConvertFromUtf32(65282)), '"');
        }

        /// <summary>
        /// </summary>
        static void cout(params object[] objs)
        {
            string output = "";

            for (int i = 0; i < objs.Length; i++)
            {
                if (objs[i] == null) objs[i] = "";

                string delim = " : ";

                if (i == objs.Length - 1) delim = "";

                output += string.Format("{0}{1}", objs[i], delim);
            }

            Console.Write(output + Environment.NewLine);
        }

    }
}
