using System;
using System.Text.RegularExpressions;

namespace RemovingProtectionFromExcel;

public static class Program
{
    static void Main(string[] args)
    {
        ConsoleEx.Write_header_line("Removing excel worksheet protection");
        ConsoleEx.Write_neutral_line("Select the directory and Excel files inside to unprotect their worksheets.");
        ConsoleEx.Write_neutral_line("Allowed file extensions are .xlsx and .xlsm.!\n");
        Regex xlsx_or_xlsm_regex = new(@"(.*\.xlsm)|(.*\.xlsx)", RegexOptions.IgnoreCase);

        string[] file_pathes = Directory_and_file_provider.Request_file_pathes_from_user_via_console(
            file_selector_by_its_name: v => xlsx_or_xlsm_regex.IsMatch(v));

        if (file_pathes.Length == 0)
        {
            ConsoleEx.Write_error_line("\nEND!");
            return;
        }

        ConsoleEx.Write_neutral_line("\nProcessing...");
        foreach (string file_path in file_pathes)
        {
            Excel_xml_content_operator.Remove_protection_from_all_excel_sheets(file_path);
        }

        ConsoleEx.Write_success_line("\nDONE!");
        Console.ReadLine();
    }
}
