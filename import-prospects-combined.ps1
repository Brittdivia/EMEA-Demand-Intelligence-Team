# import-prospects-combined.ps1
# Reads new CSV (10-05), writes data-prospects.js directly from C# to avoid memory issues
# XLSX step skipped: the 4GB CSV (2.7M rows) is a superset of the xlsx (881K rows)

$csvPath  = "C:\Users\I572929\OneDrive - SAP SE\2026\Campaign Insights AI\old\Prospect DL\Prospects 10-05.csv"
$outDir   = "C:\Users\I572929\campaign-calendar-site"
$outFile  = "$outDir\data-prospects.js"

# ── Inline C#: reads CSV and writes JS directly (no giant in-memory list) ─────
$csharp = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

public static class FastCsvProspects {
    static string EscJS(string s) {
        if (s == null) return "";
        return s.Replace("\\", "\\\\").Replace("\"", "\\\"")
                .Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
    }

    static string[] SplitCsvLine(string line) {
        var fields = new List<string>();
        bool inQuote = false;
        var cur = new StringBuilder();
        for (int i = 0; i < line.Length; i++) {
            char c = line[i];
            if (c == '"') {
                if (inQuote && i + 1 < line.Length && line[i + 1] == '"') {
                    cur.Append('"'); i++;
                } else {
                    inQuote = !inQuote;
                }
            } else if (c == ',' && !inQuote) {
                fields.Add(cur.ToString()); cur.Clear();
            } else {
                cur.Append(c);
            }
        }
        fields.Add(cur.ToString());
        return fields.ToArray();
    }

    static int FindCol(string[] headers, string name) {
        for (int i = 0; i < headers.Length; i++)
            if (string.Equals(headers[i].Trim().Trim('"').Trim(), name, StringComparison.OrdinalIgnoreCase)) return i;
        return -1;
    }

    static string G(string[] f, int idx) {
        return (idx >= 0 && idx < f.Length) ? f[idx] : "";
    }

    public static void ProcessAndWrite(string csvPath, string outputPath) {
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        int count = 0;

        using (var fs = File.Open(csvPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var reader = new StreamReader(fs, Encoding.UTF8))
        using (var outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
        using (var writer = new StreamWriter(outFs, new UTF8Encoding(false), 65536)) {

            string headerLine = reader.ReadLine();
            if (headerLine == null) return;
            if (headerLine.Length > 0 && headerLine[0] == '\uFEFF') headerLine = headerLine.Substring(1);

            string[] headers = SplitCsvLine(headerLine);
            int cId   = FindCol(headers, "ID");
            int cEid  = FindCol(headers, "External ID");
            int cFn   = FindCol(headers, "First Name");
            int cLn   = FindCol(headers, "Last Name");
            int cCo   = FindCol(headers, "Company");
            int cTou  = FindCol(headers, "Touched At");
            int cSc   = FindCol(headers, "Stage Changed At");
            int cCr   = FindCol(headers, "Created At");
            int cTg   = FindCol(headers, "Tags");
            int cSt   = FindCol(headers, "Stage Name");
            int cEm   = FindCol(headers, "Email");
            int cAu   = FindCol(headers, "Assigned Users");
            int cPn   = FindCol(headers, "Persona Name");
            int cAs   = FindCol(headers, "Active Sequences");
            int cFs   = FindCol(headers, "Finished Sequences");
            int cCf63 = FindCol(headers, "Custom Field 63");
            int cCf64 = FindCol(headers, "Custom Field 64");
            Console.WriteLine("  Cols: ID=" + cId + " Co=" + cCo + " Tags=" + cTg + " Stage=" + cSt);

            writer.Write("window.PROSPECT_DATA=[");
            bool first = true;
            string line;
            int lineNum = 0;

            while ((line = reader.ReadLine()) != null) {
                lineNum++;
                if (lineNum % 1000000 == 0) Console.WriteLine("  Row " + lineNum + "...");

                string[] f = SplitCsvLine(line);
                string id = G(f, cId);
                if (string.IsNullOrEmpty(id) || !seenIds.Add(id)) continue;

                string entry = "{\"id\":\"" + EscJS(id) + "\","
                    + "\"eid\":\"" + EscJS(G(f,cEid)) + "\","
                    + "\"fn\":\"" + EscJS(G(f,cFn)) + "\","
                    + "\"ln\":\"" + EscJS(G(f,cLn)) + "\","
                    + "\"co\":\"" + EscJS(G(f,cCo)) + "\","
                    + "\"touched\":\"" + EscJS(G(f,cTou)) + "\","
                    + "\"sc\":\"" + EscJS(G(f,cSc)) + "\","
                    + "\"created\":\"" + EscJS(G(f,cCr)) + "\","
                    + "\"tg\":\"" + EscJS(G(f,cTg)) + "\","
                    + "\"stage\":\"" + EscJS(G(f,cSt)) + "\","
                    + "\"email\":\"" + EscJS(G(f,cEm)) + "\","
                    + "\"au\":\"" + EscJS(G(f,cAu)) + "\","
                    + "\"pn\":\"" + EscJS(G(f,cPn)) + "\","
                    + "\"as\":\"" + EscJS(G(f,cAs)) + "\","
                    + "\"fs\":\"" + EscJS(G(f,cFs)) + "\","
                    + "\"cf63\":\"" + EscJS(G(f,cCf63)) + "\","
                    + "\"cf64\":\"" + EscJS(G(f,cCf64)) + "\"}";

                if (!first) writer.Write(',');
                writer.Write(entry);
                first = false;
                count++;
            }
            writer.Write("];");
            Console.WriteLine("  Done: " + count + " prospects from " + lineNum + " lines");
        }
    }
}
'@

Add-Type -TypeDefinition $csharp -Language CSharp
if ($?) { Write-Host "C# compiled OK" } else { Write-Host "C# compile FAILED"; exit 1 }

# ── Copy CSV and process ──────────────────────────────────────────────────────
Write-Host "Copying CSV to temp..."
$csvTemp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "prosp_import_tmp.csv")
[System.IO.File]::Copy($csvPath, $csvTemp, $true)
Write-Host "Copied: $([math]::Round((Get-Item $csvTemp).Length/1MB,0)) MB. Processing..."

[FastCsvProspects]::ProcessAndWrite($csvTemp, $outFile)
Remove-Item $csvTemp -Force -ErrorAction SilentlyContinue

Write-Host "Output: $([math]::Round((Get-Item $outFile).Length/1MB,1)) MB"
