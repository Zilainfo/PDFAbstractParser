using ConsoleApp.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace ConsoleApp
{
    public class Parser
    {
        public string fromPath { get; }
        public string InPath { get; }
        public int starPage { get; }
        public List<Topic> topics { get; } = new List<Topic>();

        public Parser(IUserInput input, int _starPage)
        {
            starPage = _starPage;
            Console.Write("Enter a From file directory:");
            fromPath = input.GetFromPath();
            Console.Write("Enter a In file directory:");
            InPath = input.GetInPath();
            Console.WriteLine($"Execute from  {fromPath} to {InPath}");
        }

        public int FreeCellByColName(ExcelWorksheet ws, string Name)
        {
            int index = 1;
            bool find = false;
            int i = 1;
            while (!find)
            {
                find = ws.Cells[Name + i].Value == null;
                i++;
                index = i;
            }

            return index;
        }
        public void ParseToExel()
        {
            FileInfo fi = new FileInfo(InPath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {

                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];

                int startInd = FreeCellByColName(firstWorksheet, "A") - 1;

                foreach (Topic topic in topics)
                {
                    foreach (Participant part in topic.Participants)
                    {
                        firstWorksheet.Cells["A" + startInd].Value = part.Name;
                        firstWorksheet.Cells["B" + startInd].Value = part.AffiliationNames;
                        firstWorksheet.Cells["C" + startInd].Value = part.PersonsLocation;
                        firstWorksheet.Cells["D" + startInd].Value = topic.SessionName;
                        firstWorksheet.Cells["E" + startInd].Value = topic.Title;
                        firstWorksheet.Cells["F" + startInd].Value = topic.PresentationAbstract;

                        startInd++;
                    }
                }
                //Save your file
                excelPackage.Save();
            }
        }

        public void ParseFromTopic(StringBuilder topicString)
        {
            string pattern = @"(?<session_name>^P.+TimesNewRomanPS-BoldItalicMT 9,5[\s\S]+?)(?<topic_title>^.+TimesNewRomanPS-BoldMT 9$[\s\S]+?)(?<names>^.+TimesNewRomanPS-ItalicMT 8$[\s\S]+|^.+TimesNewRomanPS-ItalicMT 9$[\s\S]+?)(?<Affiliation>(?:<\/t>$|^.+TimesNewRomanPS-ItalicMT 8$)[\s\S]+?)(?<topic>(^.+(?:TimesNewRomanPS-ItalicMT 9$|TimesNewRomanPSMT 9$)[\s\S]+))";

            RegexOptions options = RegexOptions.Multiline;
            Regex reg = new Regex(pattern, options);
            try
            {
                Match m = reg.Match(topicString.ToString()); while (m.Success)
                {
                    Topic topic = new Topic();

                    topic.SessionName = PostprocessingText(m.Groups["session_name"].Value);
                    topic.Title = PostprocessingText(m.Groups["topic_title"].Value);
                    topic.PresentationAbstract = PostprocessingText(m.Groups["topic"].Value);



                    string[] splitedMatches = new string[2];


                    if (Regex.IsMatch(PostprocessingText(m.Groups["names"].Value), "<"))
                    {
                        splitedMatches = PostprocessingText(m.Groups["names"].Value).Split('<');

                    }
                    else
                    {
                        splitedMatches[0] = PostprocessingText(m.Groups["names"].Value);
                        splitedMatches[1] = PostprocessingText(m.Groups["Affiliation"].Value);
                    }

                    bool indexOnEnd = Regex.IsMatch(splitedMatches[1], @"([\s\S]+?\d+?)(?:,|$)", options);
                    bool withIndex = Regex.IsMatch(splitedMatches[1], @"(\d+?)", options);
                    MatchCollection affiliations;
                    MatchCollection names;
                    //if Index on start
                    if (!indexOnEnd && withIndex)
                    {
                        //Add new lane
                        splitedMatches[1] = Regex.Replace("1" + splitedMatches[1], @"(\d+[\s\S]+?)", "\n$1") + "\n";
                        splitedMatches[1] += "\n";
                        splitedMatches[1] = splitedMatches[1];
                        affiliations = Regex.Matches(splitedMatches[1], @"(?<index>\d+)(?<afilation>[\s\S]+?)(,|$)(?<place>[\s\S]+?)(\n|$)", options);
                        names = Regex.Matches(splitedMatches[0], @"(?<name>[\s\S]+?)(?<index>[,\d]+)(?:,|$)", options);

                        foreach (Match name in names)
                        {
                            Match res = affiliations[0];
                            Participant participant = new Participant();

                            if (Regex.IsMatch(name.Groups["name"].Value, @","))
                            {
                                MatchCollection grpNames = Regex.Matches(name.Groups["name"].Value, @"([\s\S] +?)", options);
                                MatchCollection grpIndex = Regex.Matches(name.Groups["index"].Value, @"([\s\S] +?)", options);
                                int i = 0;
                                foreach (Match grpInd in grpIndex)
                                {
                                    try
                                    {
                                        res = (from match in affiliations
                                               where match.Groups["index"].Value.Equals(grpInd.Groups["index"].Value)
                                               select match).Single();
                                    }
                                    catch
                                    {

                                    }

                                    participant.Name = grpNames[i++].Groups["name"].Value;
                                    participant.PersonsLocation = res.Groups["place"].Value.TrimEnd().TrimEnd(',').Replace("and", "");
                                    participant.AffiliationNames = res.Groups["afilation"].Value;
                                    topic.Participants.Add(participant);
                                }
                            }
                            else
                            {

                                try
                                {
                                    res = (from match in affiliations
                                           where match.Groups["index"].Value.Equals(name.Groups["index"].Value)
                                           select match).Single();
                                }
                                catch
                                {

                                }

                                participant.Name = name.Groups["name"].Value;
                                participant.PersonsLocation = res.Groups["place"].Value.TrimEnd().TrimEnd(',').Replace("and", "");
                                participant.AffiliationNames = res.Groups["afilation"].Value;
                                topic.Participants.Add(participant);
                            }
                        }

                    }
                    else if (indexOnEnd && withIndex)
                    {
                        //Add new lane
                        /*    splitedMatches[1] = Regex.Replace(splitedMatches[1], @"(\d+?)(?:,|$)", "\n$1");*/
                        affiliations = Regex.Matches(splitedMatches[1], @"(?<afilation>[\s\S]+?)(?<index>[,\d]+)(?:,|$)", options);
                        names = Regex.Matches(splitedMatches[0], @"(?<name>[\s\S]+?)(?<index>[,\d]+)(?:,|$)", options);

                        foreach (Match name in names)
                        {
                            Participant participant = new Participant();
                            Match res = affiliations[0];

                            try
                            {
                                res = (from match in affiliations
                                       where match.Groups["index"].Value.Equals(name.Groups["index"].Value)
                                       select match).Single();
                            }
                            catch
                            {

                            }

                            participant.Name = name.Groups["name"].Value;
                            participant.PersonsLocation = res.Groups["place"].Value.TrimEnd().TrimEnd(',').Replace("and", "");
                            participant.AffiliationNames = res.Groups["afilation"].Value;
                            topic.Participants.Add(participant);
                        }

                    }
                    else
                    {
                        affiliations = Regex.Matches(splitedMatches[1], @"(?<Name>\d[A-Z][\s\S]*?,(?<place>[\s\S]*?))\n", options);


                        foreach (string name in splitedMatches[0].Split(','))
                        {
                            Participant participant = new Participant();
                            var r = Regex.Match(splitedMatches[1], @"(?<afilation>[\s\S]+?),(?<place>(?:[\s\S]+?)|)(?:$|\n)", options);
                            if (r.Success)
                            {
                                participant.Name = name;
                                participant.PersonsLocation = r.Groups["place"].Value;
                                participant.AffiliationNames = r.Groups["afilation"].Value;
                                topic.Participants.Add(participant);
                            }
                            else
                            {

                                participant.Name = name;
                                participant.AffiliationNames = splitedMatches[1];
                                topic.Participants.Add(participant);
                            }
                        }


                    }

                    topics.Add(topic);

                    m = m.NextMatch();
                }
            }
            catch (Exception)
            {

                throw;
            }

        }

        public void GetText()
        {
            bool isSessionStylePaternStart = false;
            bool isSession = false;
            Letter previos = null;

            StringBuilder sessionFullTopic = new StringBuilder();

            using (var document = PdfDocument.Open(fromPath))
            {
                for (int i = starPage; i <= document.NumberOfPages; i++)
                {
                    var page = document.GetPage(i);

                    foreach (var letter in page.Letters)
                    {
                        if (Regex.IsMatch(letter.ToString(), @"TimesNewRomanPS-BoldItalicMT 9,5"))
                        {
                            //check if Topic end
                            if (isSession)
                            {
                                this.ParseFromTopic(sessionFullTopic);
                                sessionFullTopic.Clear();
                                isSession = false;
                            }

                            if (previos != null && letter.FontName != previos.FontName && letter.PointSize != previos.PointSize)
                            {
                                sessionFullTopic.Clear();
                                isSessionStylePaternStart = true;
                            }
                        }

                        if (isSessionStylePaternStart && !(Regex.IsMatch(letter.ToString(), @"TimesNewRomanPS-BoldItalicMT 9,5")))
                        {
                            isSession = Regex.IsMatch(PostprocessingText(sessionFullTopic.ToString()), @"^P\d+$");
                            isSessionStylePaternStart = false;
                        }

                        if ((Regex.IsMatch(letter.ToString(), @"TimesNewRomanPS-ItalicMT 4,66")))
                        {

                            if (previos != null && (previos.Location.Y - letter.Location.Y) / (previos.Location.Y + letter.Location.Y) == 0.015597889626094917)
                            {
                                sessionFullTopic.Append("</t>\n");
                                previos = letter;
                            }
                        }

                        if (previos != null && (previos.Location.Y - letter.Location.Y) == 11.0)
                        {
                            sessionFullTopic.Append("</t>\n");
                            previos = letter;
                        }
                        else
                        {
                            sessionFullTopic.Append(letter.ToString() + "\n");
                            previos = letter;
                        }
                    }
                }
            }
            ParseToExel();

        }

        public string PostprocessingText(string data)
        {
            StringBuilder stringBuilder = new StringBuilder();

            var stringMatc = Regex.Matches(data, @"^(.)", RegexOptions.Multiline);
            foreach (Match m in stringMatc)
            {
                stringBuilder.Append(m.Value);
            }
            return Regex.Replace(stringBuilder.ToString(), @"(\d+www\.[\s\S]+?2018)", "");
        }




    }
}




