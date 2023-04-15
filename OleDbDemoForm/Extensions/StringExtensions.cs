using System.Text.RegularExpressions;

namespace OleDbDemoForm.Extensions;

public static class StringExtensions
{
    /// <summary>
    /// Given a string with upper and lower cased letters separate them before each upper cased characters
    /// </summary>
    /// <param name="sender">String to work against</param>
    /// <returns>String with spaces between upper-case letters</returns>
    public static string SplitCamelCase(this string sender) =>
        string.Join(" ", Regex.Matches(sender, @"([A-Z][a-z]+)")
            .Select(m => m.Value));


}