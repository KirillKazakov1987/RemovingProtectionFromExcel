using System;

namespace RemovingProtectionFromExcel;

public static class ConsoleEx
{
    public const ConsoleColor header_color = ConsoleColor.Yellow;
    public const ConsoleColor error_color = ConsoleColor.Red;
    public const ConsoleColor warning_color = ConsoleColor.DarkYellow;
    public const ConsoleColor success_color = ConsoleColor.Green;
    public const ConsoleColor neutral_color = ConsoleColor.White;
    public const ConsoleColor path_color = ConsoleColor.Magenta;
    public const ConsoleColor number_color = ConsoleColor.Cyan;

    public static void Write_header_line(object? value) => WriteLine(value, header_color);
    public static void Write_error_line(object? value) => WriteLine(value, error_color);
    public static void Write_warning_line(object? value) => WriteLine(value, warning_color);
    public static void Write_success_line(object? value) => WriteLine(value, success_color);
    public static void Write_neutral_line(object? value) => WriteLine(value, neutral_color);
    public static void Write_path_line(object? value) => WriteLine(value, path_color);
    public static void Write_number_line(object? value) => WriteLine(value, number_color);


    public static void Write_header(object? value) => Write(value, header_color);
    public static void Write_error(object? value) => Write(value, error_color);
    public static void Write_warning(object? value) => Write(value, warning_color);
    public static void Write_success(object? value) => Write(value, success_color);
    public static void Write_neutral(object? value) => Write(value, neutral_color);
    public static void Write_path(object? value) => Write(value, path_color);
    public static void Write_number(object? value) => Write(value, number_color);





    public static void WriteLine(object? value, ConsoleColor forecolor)
    {
        Console.ForegroundColor = forecolor;
        Console.WriteLine(value);
        Console.ResetColor();
    }

    public static void Write(object? value, ConsoleColor forecolor)
    {
        Console.ForegroundColor = forecolor;
        Console.Write(value);
        Console.ResetColor();
    }
}
