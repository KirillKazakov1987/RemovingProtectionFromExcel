using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace RemovingProtectionFromExcel;

public static class Directory_and_file_provider
{
    public static string[] Request_file_pathes_from_user_via_console(Predicate<string> file_selector_by_its_name)
    {
        string directory_path = Request_directory_path_from_user_via_console();
        if (directory_path.Length == 0) return Array.Empty<string>();


        DirectoryInfo di = new(directory_path);

        FileInfo[] all_files = di.GetFiles("*.*", SearchOption.TopDirectoryOnly);

        FileInfo[] pertinent_files = all_files
            .Where(v => file_selector_by_its_name(v.Name))
            .OrderBy(f => f.Name)
            .ToArray();

        ConsoleEx.Write_neutral("Count all files in the directory = ");
        ConsoleEx.Write_number_line(all_files.Length);
        ConsoleEx.Write_neutral("Count pertinent files in the directory = ");
        ConsoleEx.Write_number_line(pertinent_files.Length);


        if (all_files.Length == 0)
        {
            ConsoleEx.Write_error_line("There are no files.");
            return Array.Empty<string>();
        }

        if (pertinent_files.Length == 0)
        {
            ConsoleEx.Write_error_line("There are no pertinent files.");
            return Array.Empty<string>();
        }

        if (pertinent_files.Length == 1)
        {
            ConsoleEx.Write_success("Selected 1 file: \'");
            ConsoleEx.Write_path(pertinent_files[0].Name);
            ConsoleEx.Write_success_line("\'");
            return new string[] { pertinent_files[0].FullName };
        }


        ConsoleEx.Write_header_line("\nFILES:");
        ConsoleEx.Write_neutral("Index".PadRight(6));
        ConsoleEx.Write_neutral("Name");
        Console.WriteLine();

        for (int i = 0; i < pertinent_files.Length; i++)
        {
            ConsoleEx.Write_number((i + 1).ToString().PadRight(6));
            ConsoleEx.Write_path_line(pertinent_files[i].Name);
        }

        ConsoleEx.Write_neutral_line("Please set index or range of indexes or space separated list of index from table above.\nExamples:");

        ConsoleEx.Write_number("2".PadRight(10));
        ConsoleEx.Write_neutral_line("→ select file with index 2");

        ConsoleEx.Write_number("3-5".PadRight(10));
        ConsoleEx.Write_neutral_line("→ select 3 files with indexes 3, 4 and 5");

        ConsoleEx.Write_number("2 4 6 7".PadRight(10));
        ConsoleEx.Write_neutral_line("→ select 4 files with indexes 2, 4, 6 and 7");
        ConsoleEx.Write_neutral_line("Set index:");

        while (true)
        {
            string? user_input = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(user_input)) continue;

            int[] onedexes = User_input_to_integer_array(user_input, pertinent_files.Length);

            if (onedexes.Length == 0)
            {
                ConsoleEx.Write_error_line("There are no pertinent indexes in input");
                return Array.Empty<string>();
            }

            ConsoleEx.Write_success($"Selected file indexes ({onedexes.Length}): ");
            ConsoleEx.Write_number_line(string.Join(' ', onedexes));

            return onedexes.Select(i => pertinent_files[i-1].FullName).ToArray();
        }
    }



    private static int[] User_input_to_integer_array(
        string user_input,
        int max_inclusive_integer)
    {
        Regex single_integer_regex = new(@"\d");
        Regex range_regex = new(@"\d-\d");
        Regex list_regex = new(@"\d\s+\d");

        if (range_regex.IsMatch(user_input))
        {
            string[] entries = user_input.Split('-', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
            if (entries.Length != 2) goto UNEXPECTED_USER_INPUT;

            if (int.TryParse(entries[0], out int begin) == false) goto UNEXPECTED_USER_INPUT;
            if (int.TryParse(entries[1], out int end) == false) goto UNEXPECTED_USER_INPUT;

            begin = Math.Max(begin, 1);
            end = Math.Min(end, max_inclusive_integer);

            if (end < begin) goto UNEXPECTED_USER_INPUT;

            return Enumerable.Range(begin, end - begin + 1).ToArray();
        }

        if (list_regex.IsMatch(user_input))
        {
            string[] entries = user_input.Split(' ', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

            List<int> integers = new(entries.Length);

            for (int i = 0; i < entries.Length; i++)
            {
                if (int.TryParse(entries[i], out int value) == false) goto UNEXPECTED_USER_INPUT;

                if (value <= 0) goto UNEXPECTED_USER_INPUT;

                if (value <= max_inclusive_integer) integers.Add(value);
            }

            return integers.Distinct().ToArray();
        }

        if (single_integer_regex.IsMatch(user_input))
        {
            if (int.TryParse(user_input, out int value) == false) goto UNEXPECTED_USER_INPUT;

            if (value <= 0) goto UNEXPECTED_USER_INPUT;

            if (value <= max_inclusive_integer) return new int[] { value };
            else return Array.Empty<int>();
        }


        UNEXPECTED_USER_INPUT:
        ConsoleEx.Write_error($"Unexpected used input: \'");
        ConsoleEx.Write_number(user_input);
        ConsoleEx.Write_error_line($"\'");
        return Array.Empty<int>();
    }



    private class DirectoryInfoEqualityComparer : IEqualityComparer<DirectoryInfo>
    {
        public static readonly DirectoryInfoEqualityComparer Instance = new();

        public bool Equals(DirectoryInfo? x, DirectoryInfo? y)
        {
            if (x != null && y != null)
            {
                return x.FullName.Equals(y.FullName);
            }
            else
            {
                return object.Equals(x, y);
            }
        }

        public int GetHashCode([DisallowNull] DirectoryInfo obj)
        {
            return obj.FullName.GetHashCode();
        }
    }


    public static string Request_directory_path_from_user_via_console()
    {
        List<DirectoryInfo> L = new();


        L.Add(new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)));
        L.Add(new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)));
        L.Add(new DirectoryInfo(Environment.CurrentDirectory));

        DirectoryInfo? p;

        for (int i = 0; i < 10; i++)
        {
            p = L[^1].Parent; if (p != null) L.Add(p);
        }


        const string previous_selected_directory_holder = "previous_selected_directory_holder.txt";
        string previous_selected_directory_holder_path = Path.Combine(Environment.CurrentDirectory, previous_selected_directory_holder);
        
        if (File.Exists(previous_selected_directory_holder_path))
        {
            using (StreamReader sr = new(previous_selected_directory_holder_path))
            {
                string? prev_path = sr.ReadLine();

                if (string.IsNullOrEmpty(prev_path) == false)
                {
                    L.Add(new DirectoryInfo(prev_path));
                }
            }
        }



        L = L.Distinct(DirectoryInfoEqualityComparer.Instance).Where(v => v.Parent != null).ToList();


        ConsoleEx.Write_header_line("DIRECTORIES:");
        ConsoleEx.Write_neutral("Index".PadRight(6));
        ConsoleEx.Write_neutral("Path");
        Console.WriteLine();

        for (int i = 0; i < L.Count; i++)
        {
            ConsoleEx.Write_number((i + 1).ToString().PadRight(6));
            ConsoleEx.Write_path_line(L[i]);
        }


        ConsoleEx.Write_neutral("Please set");
        ConsoleEx.Write_path(" directory path ");
        ConsoleEx.Write_neutral("or");
        ConsoleEx.Write_number(" directory index ");
        ConsoleEx.Write_neutral("from table above and press ENTER:");
        Console.WriteLine();


        string result;

        while (true)
        {
            string? user_input = Console.ReadLine()?.Trim();

            if (string.IsNullOrEmpty(user_input)) continue;

            
            if (int.TryParse(user_input, out int directory_onedex)) 
            {
                int directory_index = directory_onedex - 1;
                if (directory_index < 0 || directory_index >= L.Count)
                {
                    ConsoleEx.Write_error_line($"Index is too small or too big: {directory_onedex}.");
                    result = string.Empty;
                    return result;
                }
                else
                {
                    result = L[directory_index].FullName;
                    break;
                }
            }

            DirectoryInfo di = new(user_input);

            if (di.Exists)
            {
                result = di.FullName;
                break;
            }
            else
            {
                ConsoleEx.Write_error($"Directory path \'");
                ConsoleEx.Write_path(user_input);
                ConsoleEx.Write_error($"\' is not valid or does not exist: \'");
                Console.WriteLine();
                result = string.Empty;
                return result;
            }
        }

        ConsoleEx.Write_success($"Selected directory path: \'");
        ConsoleEx.Write_path(result);
        ConsoleEx.Write_success("\'");
        Console.WriteLine();

        try
        {
            File.WriteAllText(previous_selected_directory_holder_path, result);
        }
        catch { }

        return result;
    }


}