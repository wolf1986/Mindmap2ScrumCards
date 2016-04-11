//css_ref System.Xml;
//css_ref Office;
//css_ref NDesk.Options;

//Need office Interop, maybe in: C:\Program Files\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14\Office.dll;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

using System.Drawing;
using NDesk.Options;

class Script
{
    // Priority Colors
    public static Dictionary<int, Color> TranslationsColorPriority = new Dictionary<int, Color>
        {
            {1, Color.OrangeRed},
            {2, Color.Yellow},
            {3, Color.Orange},
            {4, Color.LightGreen},
            {5, Color.DodgerBlue},
            {6, Color.Gray},
        };

    // Get number in braces in the end of the string - Some User Story (12) 
    static public Tuple<string, string> SplitBracesStr(string name)
    {
        var ret1 = name;
        var ret2 = "";

        try
        {
            var match = Regex.Match(name, @"(?<Text>.*?)\s*\((?<Bracets>\d*)\)\s*$");

            if (match.Success)
            {
                var groups = match.Groups;
                ret1 = groups["Text"].Value;
                ret2 = groups["Bracets"].Value;
            }
        }
        catch { }    // Ignore errors

        return new Tuple<string, string>(ret1, ret2);
    }

    // Transform string with braces to CSV notation
    static public string TransformSplitBraces(string name)
    {
        var tup = SplitBracesStr(name);
        return tup.Item1 + ", " + tup.Item2;
    }

    static public string TransformTrimBraces(string name)
    {
        var tup = SplitBracesStr(name);
        return tup.Item1;
    }

    public static IEnumerable<string> GetListElements(IEnumerable<XElement> elements, string attributeName, Func<string, string> transformation = null)
    {
        if (transformation == null)
            transformation = (x => x);

        return
            (
                from element in elements
                select transformation(element.Attribute(attributeName).Value)
            ).Except(new[] { null, "" });   // Skip Null & Empty tasks
    }

    static public string GetElementsStr(IEnumerable<XElement> elements, string attributeName, string linePrefix = "\n\t", Func<string, string> transformation = null)
    {
        return linePrefix + string.Join(linePrefix, GetListElements(elements, attributeName, transformation));
    }

    static public void PrintListElements(IEnumerable<XElement> elements, string attributeName, string linePrefix = "\n\t", Func<string, string> transformation = null)
    {
        var str = GetElementsStr(elements, attributeName, linePrefix, transformation);
        Console.WriteLine(str);
    }

    static public void AddNewSlide(
        Presentation presentation,
        string pathFileTemplate,
        Dictionary<string, string> translationTitles,
        Dictionary<string, Color> translationColors)
    {
        const int cThresholdDarkColor = 500;

        // Copy a slide from the template file
        presentation.Slides.InsertFromFile(pathFileTemplate, presentation.Slides.Count, 1, 1);

        // Modify shapes in the new slide
        var slide = presentation.Slides[presentation.Slides.Count];
        var shapes_all = slide.Shapes.Cast<Shape>();
        var shapes_our =
            from shape in shapes_all
            where
                shape.HasTextFrame == MsoTriState.msoTrue &&
                shape.TextFrame.HasText == MsoTriState.msoTrue &&
                shape.TextFrame.TextRange.Text.StartsWith("$")
            select shape;

        var not_found_keys = new List<string>();

        // Set the requested params
        foreach (var key in translationTitles.Keys)
        {
            var shape = shapes_all.FirstOrDefault(
                x =>
                    x.HasTextFrame == MsoTriState.msoTrue &&
                    x.TextFrame.HasText == MsoTriState.msoTrue &&
                    x.TextFrame.TextRange.Text == key
            );

            if (null == shape)
            {
                not_found_keys.Add(key);
                continue;
            }

            shape.TextFrame.TextRange.Text = translationTitles[key];

            // Delete "-1" priorities
            if (key == "$P")
            {
                var priority = int.Parse(translationTitles[key]);

                if (translationTitles[key] == "-1")
                    shape.Delete();
                else
                {
                    var fore_color = TranslationsColorPriority[priority];
                    shape.Fill.ForeColor.RGB = ToRgb(fore_color);

                    // Light text for dark back-grounds
                    if (fore_color.R + fore_color.G + fore_color.B < cThresholdDarkColor)
                        shape.TextFrame.TextRange.Font.Color.RGB = ToRgb(Color.White);
                    else
                        shape.TextFrame.TextRange.Font.Color.RGB = ToRgb(Color.Black);
                }
            }

            if (translationColors.ContainsKey(key))
            {
                var fore_color = translationColors[key];
                shape.Fill.ForeColor.RGB = ToRgb(fore_color);

                shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;

                // Light text for dark back-grounds
                if (fore_color.R + fore_color.G + fore_color.B < cThresholdDarkColor)
                {
                    shape.TextFrame.TextRange.Font.Color.RGB = ToRgb(Color.White);
                }
            }
        }

        // Search inside tables
        foreach (var key in not_found_keys)
        {
            var shape = shapes_all.FirstOrDefault(
                x =>
                    x.HasTable == MsoTriState.msoTrue
            );

            for (var i_row = 1; i_row <= shape.Table.Rows.Count; i_row++)
                for (var i_col = 1; i_col <= shape.Table.Rows[i_row].Cells.Count; i_col++)
                {
                    var cell = shape.Table.Rows[i_row].Cells[i_col];

                    if (cell.Shape.HasTextFrame == MsoTriState.msoTrue &&
                        cell.Shape.TextFrame.TextRange.Text == key)
                    {
                        cell.Shape.TextFrame.TextRange.Text = translationTitles[key];
                    }
                }
        }
    }

    static public int ToRgb(Color c)
    {
        return (((int)c.B << 16) + ((int)c.G << 8) + (int)c.R);
    }

    static public string ExtractFirstIcon(XElement element)
    {
        // Find first  icon, ignore if failed
        try
        {
            var icon_str = element.Element("icon").Attribute("BUILTIN").Value;
            return icon_str;
        }
        catch
        {
            return "";
        }
    }

    static public string IconStrToPriority(string str)
    {
        try
        {
			var match = Regex.Match(str, @"full-(\d)$");
            if (match.Success)
				return match.Groups[1].Value;
        }
        catch {}
		return "-1";
    }

    static public void Main(string[] args)
    {
        var path_mindmap = "";
        var path_template = "";
        var help = false;
		
		// Debug
		// -m planning.mm -t "Tasks - Template - No Left.pptx"
		path_mindmap = "planning.mm";
		path_template = "Tasks - Template - No Left.pptx";		

		/*
        var p = new OptionSet() {
            { "m=|path-mindmap=", v => path_mindmap = v },
            { "t=|path-template=", v => path_template = v },
            { "h|help", v => help = true}
          };

        p.Parse(args);

        // Print usage if needed, stop execution
        if (help || string.IsNullOrEmpty(path_mindmap) || string.IsNullOrEmpty(path_template))
        {
            p.WriteOptionDescriptions(Console.Out);
            return;
        }
		*/

        // Transform to full paths
        path_template = Path.GetFullPath(path_template);
        path_mindmap = Path.GetFullPath(path_mindmap);

        //var path_template =  
        //    Path.Combine(Directory.GetCurrentDirectory(), "Tasks - Template.pptx");

        var path_file_save_as =
            Path.Combine(Path.GetDirectoryName(path_mindmap), Path.GetFileNameWithoutExtension(path_mindmap) + ".pptx");

        File.Copy(path_template, path_file_save_as, true);

        // Define 13 custom colors that look good on notes
        var list_colors_to_use = new Queue<Color>(new[] {
            Color.PaleGreen,
            Color.Blue,
            Color.Red,
            Color.Orange,
            Color.LightGray,
            Color.Purple,
            Color.Green,
            Color.LightGoldenrodYellow,
            Color.LightSkyBlue,
            Color.LightPink,
            Color.Black,
            Color.DarkRed,
            Color.Orchid,
        });

        // Create a new powerpoint presentation 
        var pp_app = new Application
        {
            WindowState = PpWindowState.ppWindowMinimized
        };

        // Copy the original file in order to take also screen aspect ratio, delete the first slide

        var pp_pres = pp_app.Presentations.Open(path_file_save_as);
        pp_pres.Slides[1].Delete();

        var xml_mm = XDocument.Load(path_mindmap);

        // Find all user-stories
        var elements_stories = xml_mm.Element("map").Elements().First().Elements("node");

        // Output a csv-style list
        Console.WriteLine("* List of User-Stories:");
        PrintListElements(elements_stories, "TEXT", "\n\t", TransformTrimBraces);
        Console.WriteLine("\n--- --- --- ");

        // Output all stories information in a tree
        foreach (var element_story in elements_stories)
        {
            var tup_story_name =
                SplitBracesStr(element_story.Attribute("TEXT").Value.Trim(new[] { ' ', '\t' }));

            var story_name = tup_story_name.Item1;
            Console.WriteLine(story_name);

            var elements_tasks = element_story.Elements("node");

            var story_priority = IconStrToPriority(ExtractFirstIcon(element_story));

            // Pick a color for the user-story
            var color = list_colors_to_use.Dequeue();

            // Output a csv-style list
            Console.WriteLine("* List of tasks:");
            PrintListElements(elements_tasks, "TEXT", "\n", TransformSplitBraces);
            Console.WriteLine();

            // Find all tasks
            foreach (var element_task in elements_tasks)
            {
                var tup_task_name = SplitBracesStr(element_task.Attribute("TEXT").Value);
                var task_name = tup_task_name.Item1;
                var task_estimation = tup_task_name.Item2;
                //Console.WriteLine("\t\n" + task_name);

                // Find first priority icon
                var task_priority = IconStrToPriority(ExtractFirstIcon(element_task));

                if (task_priority == "-1" && story_priority != "-1")
                    task_priority = story_priority;

                // Find all task-comments
                var elements_comments = element_task.Elements("node");

                // Dont spam log with task comments
                //PrintListElements(elements_comments, "TEXT", "\n\t\t");
                //Console.WriteLine();

                var translation_titles = new Dictionary<string, string>
                    {
                        {"$UserStory", story_name},
                        {"$TaskTitle", task_name},
                        {"$Comments", GetElementsStr(elements_comments,"TEXT","\n").Trim()},
                        {"$Left", task_estimation},
                        {"$P", task_priority},
                    };

                var translation_colors = new Dictionary<string, Color>
                    {
                        {"$UserStory", color},
                    };

                AddNewSlide(
                    pp_pres,
                    path_template,
                    translation_titles,
                    translation_colors
                );
            }
            Console.WriteLine("--- --- --- ");
        }

        // Save and close the presentation
        pp_pres.Save();
        pp_pres.Close();
        pp_app.Quit();
    }
}