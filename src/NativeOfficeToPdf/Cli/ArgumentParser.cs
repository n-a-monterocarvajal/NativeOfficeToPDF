namespace NativeOfficeToPdf.Cli;

internal enum Verb
{
    Convert,
    Install,
    Uninstall,
    CheckUpdates,
    Version,
    Help,
}

/// <param name="Error">Nulo si la línea de comandos es válida; si no, el mensaje para el usuario.</param>
internal sealed record ParsedCommand(
    Verb Verb,
    string? Source = null,
    string? Destination = null,
    bool Overwrite = false,
    bool Open = false,
    bool Quiet = false,
    string? Error = null)
{
    public bool IsValid => Error is null;

    public static ParsedCommand Invalid(string error) => new(Verb.Help, Error: error);
}

internal static class ArgumentParser
{
    public static ParsedCommand Parse(string[] arguments)
    {
        if (arguments.Length == 0)
        {
            return ParsedCommand.Invalid("Falta el archivo de origen.");
        }

        Verb verb = Verb.Convert;
        int index = 0;

        // El primer token puede ser un verbo, pero un archivo que se llame igual que un verbo gana:
        // el uso principal es «NativeOfficeToPdf.exe <archivo>», y no queremos que un documento
        // llamado "install.docx"… ni uno llamado "install" a secas… se interprete como orden.
        if (TryParseVerb(arguments[0], out Verb parsedVerb) && !File.Exists(arguments[0]))
        {
            verb = parsedVerb;
            index = 1;
        }

        string? source = null;
        string? destination = null;
        bool overwrite = false;
        bool open = false;
        bool quiet = false;

        for (; index < arguments.Length; index++)
        {
            string argument = arguments[index];

            if (IsFlag(argument, "overwrite", "o"))
            {
                overwrite = true;
            }
            else if (IsFlag(argument, "open"))
            {
                open = true;
            }
            else if (IsFlag(argument, "quiet", "q"))
            {
                quiet = true;
            }
            else if (argument.StartsWith('-') || argument.StartsWith('/'))
            {
                return ParsedCommand.Invalid($"Opción desconocida: {argument}");
            }
            else if (source is null)
            {
                source = argument;
            }
            else if (destination is null)
            {
                destination = argument;
            }
            else
            {
                return ParsedCommand.Invalid($"Sobra un argumento: {argument}");
            }
        }

        if (verb == Verb.Convert && string.IsNullOrWhiteSpace(source))
        {
            return ParsedCommand.Invalid("Falta el archivo de origen.");
        }

        if (verb != Verb.Convert && source is not null)
        {
            return ParsedCommand.Invalid($"El verbo '{arguments[0]}' no acepta argumentos: {source}");
        }

        return new ParsedCommand(verb, source, destination, overwrite, open, quiet);
    }

    private static bool TryParseVerb(string token, out Verb verb)
    {
        switch (token.ToLowerInvariant())
        {
            case "convert":
            case "convertir":
                verb = Verb.Convert;
                return true;
            case "install":
            case "instalar":
                verb = Verb.Install;
                return true;
            case "uninstall":
            case "desinstalar":
                verb = Verb.Uninstall;
                return true;
            case "check-updates":
            case "--check-updates":
                verb = Verb.CheckUpdates;
                return true;
            case "-v":
            case "--version":
                verb = Verb.Version;
                return true;
            case "-h":
            case "-?":
            case "/?":
            case "--help":
            case "help":
                verb = Verb.Help;
                return true;
            default:
                verb = Verb.Convert;
                return false;
        }
    }

    private static bool IsFlag(string argument, string name, string? shortName = null)
    {
        if (argument.Equals("--" + name, StringComparison.OrdinalIgnoreCase) ||
            argument.Equals("/" + name, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return shortName is not null &&
            argument.Equals("-" + shortName, StringComparison.Ordinal);
    }
}
