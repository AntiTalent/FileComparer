using System.Xml.Linq;

namespace ExcelComparer;

/// <summary>
///   Compare XML files
/// </summary>
/// <remarks>
///   - Find tags not present in file2
///   - Find missing or different attributes on matching tags
/// </remarks>
public class ComparerXml : IFileComparer
{
    private Dictionary<string, string> _namespaces;

    public void Compare(string file1, string file2)
    {
        var doc1 = XDocument.Load(file1);
        var doc2 = XDocument.Load(file2);
        var root1 = doc1.Root;
        var root2 = doc2.Root;

        if (doc1.Declaration?.ToString() != doc2.Declaration?.ToString())
            Console.WriteLine("Declarations does not match");

        if (root1 == null || root2 == null) return;

        _namespaces = root1
            .Attributes()
            .Where(a => a.IsNamespaceDeclaration)
            .Skip(1)
            .ToDictionary(a => a.Value, a => a.Name.LocalName);

        CompareAttributes(root1, root2, 0);
        CompareElements(root1, root2, 0);
    }

    private void CompareElements(XElement e1, XElement e2, int indent)
    {
        var d1 = e1.Elements().ToList();
        var d2 = e2.Elements().ToList();

        var i2 = 0;
        for (var i = 0; i < d1.Count; ++i)
        {
            if (i2 < d2.Count && d1[i].Name == d2[i2].Name)
            {
                CompareAttributes(d1[i], d2[i2], indent);
                CompareElements(d1[i], d2[i2], indent + 1);

                ++i2;
            }
            else
            {
                Console.WriteLine($"{Indent(indent)}{XmlString(d1[i])}");
            }
        }
    }

    private string XmlString(XElement e) =>
        $"<{e.Name.LocalName} {string.Join(' ', e.Attributes().Select(a => $"{Name(a.Name)}=\"{a.Value}\""))}>";

    private string Name(XName n) => NamespacePrefix(n.Namespace) + n.LocalName;

    private string NamespacePrefix(XNamespace ns) =>
        _namespaces.TryGetValue(ns.NamespaceName, out var prefix) ? $"{prefix}:" : string.Empty;

    private static string Indent(int indent) => new('\t', indent);

    private void CompareAttributes(XElement e1, XElement e2, int indent)
    {
        var errors = new List<string>();
        foreach (var attribute in e1.Attributes())
        {
            var a2 = e2.Attribute(attribute.Name);
            if (a2 == null)
                errors.Add($"{Indent(indent + 1)}Missing attribute '{attribute.Name.LocalName}'");
            else if (attribute.Value != a2.Value)
                errors.Add($"{Indent(indent + 1)}Different attribute value '{attribute.Name.LocalName}' {attribute.Value} != {a2.Value}");
        }

        if (errors.Count <= 0) return;

        Console.WriteLine($"{Indent(indent)}{XmlString(e2)}");
        foreach (var error in errors)
            Console.WriteLine(error);
    }
}