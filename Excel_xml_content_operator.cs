using System;
using System.Xml;
using System.Text;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.IO;

namespace RemovingProtectionFromExcel;


public static class Excel_xml_content_operator
{
    private static readonly Regex regex_xml_tag_of_sheet_protection = new("<sheetProtection.+?/>");
    private static readonly Regex regex_full_name_of_worksheet_entry = new(@"xl/worksheets/sheet\d+\.xml");



    private static void Copy_content_from_entry_to_entry(
        in ZipArchiveEntry src_entry, 
        ZipArchiveEntry dst_entry)
    {
        Stream src_entry_stream = src_entry.Open();
        Stream dst_entry_stream = dst_entry.Open();
        src_entry_stream.CopyTo(dst_entry_stream);
        src_entry_stream.Dispose();
        dst_entry_stream.Dispose();
    }



    private static string Remove_xml_tag_of_sheet_protection(
        string sheet_xml_content)
    {
        string new_content = regex_xml_tag_of_sheet_protection.Replace(sheet_xml_content, string.Empty);
        return new_content;
    }



    private static string Stream_to_string(Stream stream)
    {
        StreamReader reader = new(stream);
        string text = reader.ReadToEnd();
        reader.Dispose();
        return text;
    }



    private static bool Check_supposed_excel_file_path(string excel_file_path)
    {
        if (File.Exists(excel_file_path) == false)
        {
            ConsoleEx.Write_error_line("File is not found. Path: ");
            ConsoleEx.Write_path_line(excel_file_path);

            return false;
        }

        string extension = Path.GetExtension(excel_file_path);
        if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) == false
            && string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase) == false)
        {
            ConsoleEx.Write_error("File extension \'");
            ConsoleEx.Write_path(extension);
            ConsoleEx.Write_error("\' is invalid.");
            ConsoleEx.Write_error("Expected extensions: ");
            ConsoleEx.Write_path(".xlsx");
            ConsoleEx.Write_error(" or ");
            ConsoleEx.Write_path(".xlsm");
            ConsoleEx.Write_error($".");
            return false;
        }

        return true;
    }



    private static void Transform_archive_entries(
        ZipArchive src_archive,
        ZipArchive dst_archive,
        Predicate<string> entry_full_name_selector_for_transformation,
        Func<string, string> string_content_transformer)
    {
        foreach (ZipArchiveEntry src_entry in src_archive.Entries)
        {
            if (entry_full_name_selector_for_transformation(src_entry.FullName))
            {
                Stream src_entry_stream = src_entry.Open();
                string old_content = Stream_to_string(src_entry_stream);
                string new_content = string_content_transformer(old_content);
                src_entry_stream.Dispose();

                ZipArchiveEntry dst_entry = dst_archive.CreateEntry(src_entry.FullName);
                Stream dst_entry_stream = dst_entry.Open();
                StreamWriter dst_entry_stream_writer = new(dst_entry_stream);
                dst_entry_stream_writer.Write(new_content);
                dst_entry_stream_writer.Dispose();
                dst_entry_stream.Dispose();
            }
            else
            {
                ZipArchiveEntry dst_entry = dst_archive.CreateEntry(src_entry.FullName, CompressionLevel.SmallestSize);
                Copy_content_from_entry_to_entry(src_entry, dst_entry);
            }
        }
    }



    private static string Transform_source_to_destination_file_path(string src_path)
    {
        if (File.Exists(src_path) == false)
        {
            throw new FileNotFoundException($"File is not found. Path = \'{src_path}\'");
        }

        string dir_path = Path.GetDirectoryName(src_path) ?? string.Empty;

        string file_name = Path.GetFileName(src_path);

        string new_file_name = "NEW " + file_name;

        string dst_path = Path.Combine(dir_path, new_file_name);

        return dst_path;
    }



    public static void Remove_protection_from_all_excel_sheets(
        string src_excel_file_path,
        Func<string, string>? src_to_dst_excel_file_path = null)
    {
        if (src_to_dst_excel_file_path == null) src_to_dst_excel_file_path = Transform_source_to_destination_file_path;

        Check_supposed_excel_file_path(src_excel_file_path);

        ZipArchive src_archive = ZipFile.Open(src_excel_file_path, ZipArchiveMode.Read);

        MemoryStream dst_archive_stream = new();
        using (ZipArchive dst_archive = new(dst_archive_stream, ZipArchiveMode.Create, true))
        {
            Transform_archive_entries(
                src_archive: src_archive,
                dst_archive: dst_archive,
                entry_full_name_selector_for_transformation: entry_full_name => regex_full_name_of_worksheet_entry.IsMatch(entry_full_name),
                string_content_transformer: content => Remove_xml_tag_of_sheet_protection(content));
        }

        src_archive.Dispose();

        string dst_file_path = src_to_dst_excel_file_path(src_excel_file_path);

        try
        {
            using (FileStream dst_file_stream = new(dst_file_path, FileMode.Create))
            {
                dst_archive_stream.Seek(0, SeekOrigin.Begin);
                dst_archive_stream.CopyTo(dst_file_stream);
            }

            ConsoleEx.Write_success("New file \'");
            ConsoleEx.Write_path(dst_file_path);
            ConsoleEx.Write_success("\' saved.");
        }
        catch (Exception ex)
        {
            ConsoleEx.Write_error_line(ex);
        }
        finally
        {
            dst_archive_stream.Dispose();
        }

        Console.WriteLine();
    }
}
