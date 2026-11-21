using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace DataHelperTaskPanePoc
{
    // Wire data is never interpreted as code, an address, or a macro name.
    public sealed class BridgeRequest
    {
        static readonly string[] PlainActions = new[] { "Unmerge", "Space", "Unit", "Unit:Force", "Date", "Mark", "BlankRows", "Dedup", "Split", "Delivery", "ClearTemporary" };
        public string Id, Command, Actions, Path, License;

        static Dictionary<string, object> Object(object value)
        {
            var result = value as Dictionary<string, object>;
            if (result == null) throw new ArgumentException("Object required");
            return result;
        }

        static void Keys(Dictionary<string, object> value, params string[] allowed)
        {
            foreach (var key in value.Keys)
                if (Array.IndexOf(allowed, key) < 0) throw new ArgumentException("Unknown field: " + key);
        }

        static string Text(Dictionary<string, object> obj, string name)
        {
            object value;
            if (!obj.TryGetValue(name, out value) || !(value is string)) throw new ArgumentException("String required: " + name);
            return (string)value;
        }

        static string ValidateSingleColLetter(string raw, string name)
        {
            var col = (raw ?? "").Trim().ToUpperInvariant();
            if (!Regex.IsMatch(col, @"\A[A-Z]{1,3}\z")) throw new ArgumentException("Invalid column letter: " + name);
            int idx = 0;
            for (int i = 0; i < col.Length; i++) idx = idx * 26 + (col[i] - 'A' + 1);
            if (idx < 1 || idx > 16384) throw new ArgumentException("Column letter out of range (A-XFD): " + name);
            return col;
        }

        static string ColLetter(Dictionary<string, object> obj, string name)
        {
            return ValidateSingleColLetter(Text(obj, name), name);
        }

        static string[] ColLetterList(Dictionary<string, object> obj, string name, int minCount, int maxCount)
        {
            var raw = Text(obj, name).Trim();
            if (raw.Length == 0) throw new ArgumentException("Invalid column letter: " + name);
            var parts = raw.Split(',');
            if (parts.Length < minCount || parts.Length > maxCount)
                throw new ArgumentException(String.Format("{0} must specify {1} to {2} columns", name, minCount, maxCount));
            var seen = new HashSet<string>();
            var normalized = new string[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                var c = ValidateSingleColLetter(parts[i], name);
                if (!seen.Add(c)) throw new ArgumentException("Duplicate column in " + name + ": " + c);
                normalized[i] = c;
            }
            return normalized;
        }

        static string ValidateSingleSheetName(string raw, string name)
        {
            var sheet = (raw ?? "").Trim();
            if (sheet.Length < 1 || sheet.Length > 31 || sheet.IndexOfAny(new[] { ':', '|', '\r', '\n', '\0', '\\', '/', '?', '*', '[', ']' }) >= 0)
                throw new ArgumentException("Invalid worksheet name: " + name);
            return sheet;
        }

        static string SheetName(Dictionary<string, object> obj, string name)
        {
            return ValidateSingleSheetName(Text(obj, name), name);
        }

        static string[] SheetNameList(Dictionary<string, object> obj, string name, int minCount, int maxCount)
        {
            var raw = Text(obj, name).Trim();
            if (raw.Length == 0) throw new ArgumentException("Invalid worksheet list: " + name);
            var parts = raw.Split(',');
            if (parts.Length < minCount || parts.Length > maxCount)
                throw new ArgumentException(String.Format("{0} must specify {1} to {2} worksheets", name, minCount, maxCount));
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var normalized = new string[parts.Length];
            for (int i = 0; i < parts.Length; i++)
            {
                var s = ValidateSingleSheetName(parts[i], name);
                if (s.IndexOf(',') >= 0) throw new ArgumentException("Worksheet name in list cannot contain comma: " + name);
                if (!seen.Add(s)) throw new ArgumentException("Duplicate worksheet in " + name + ": " + s);
                normalized[i] = s;
            }
            return normalized;
        }

        static string ValidateColMap(string raw)
        {
            var trimmed = (raw ?? "").Trim();
            if (trimmed.Length == 0) return "";
            if (trimmed.Length > 512 || trimmed.IndexOfAny(new[] { ':', '|', '\r', '\n', '\0' }) >= 0)
                throw new ArgumentException("Invalid colMap characters or length");
            var pairs = trimmed.Split(';');
            if (pairs.Length > 32) throw new ArgumentException("Too many colMap entries");
            var normalizedPairs = new List<string>();
            var seenAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in pairs)
            {
                var item = p.Trim();
                if (item.Length == 0) continue;
                int eq = item.IndexOf('=');
                if (eq <= 0 || eq >= item.Length - 1 || item.IndexOf('=', eq + 1) >= 0)
                    throw new ArgumentException("Invalid colMap pair (expected Alias=Standard): " + item);
                var left = item.Substring(0, eq).Trim();
                var right = item.Substring(eq + 1).Trim();
                if (left.Length == 0 || left.Length > 64 || right.Length == 0 || right.Length > 64)
                    throw new ArgumentException("Invalid colMap header length");
                if (!seenAliases.Add(left))
                    throw new ArgumentException("Duplicate alias in colMap: " + left);
                normalizedPairs.Add(left + "=" + right);
            }
            return String.Join(";", normalizedPairs);
        }

        static string ParseLookupSpec(Dictionary<string, object> options)
        {
            Keys(options, "sourceSheet", "sourceKeyCol", "returnCol", "targetKeyCol", "targetWriteCol", "dupRule", "writeMode");
            var srcSheet = SheetName(options, "sourceSheet");
            var srcKey = ColLetter(options, "sourceKeyCol");
            var retCols = ColLetterList(options, "returnCol", 1, 16);
            for (int i = 0; i < retCols.Length; i++)
                if (retCols[i] == srcKey) throw new ArgumentException("Source key and return column must differ");
            var tgtKey = ColLetter(options, "targetKeyCol");
            var rawWrite = Text(options, "targetWriteCol").Trim().ToUpperInvariant();
            string tgtWriteSpec;
            if (rawWrite == "AUTO")
            {
                tgtWriteSpec = "AUTO";
            }
            else
            {
                var writeCols = ColLetterList(options, "targetWriteCol", 1, 16);
                if (writeCols.Length != retCols.Length)
                    throw new ArgumentException("targetWriteCol count must match returnCol count or be AUTO");
                for (int i = 0; i < writeCols.Length; i++)
                    if (writeCols[i] == tgtKey) throw new ArgumentException("Target key and write column must differ");
                tgtWriteSpec = String.Join(",", writeCols);
            }
            var dupRule = Text(options, "dupRule");
            if (Array.IndexOf(new[] { "abort", "first", "skip" }, dupRule) < 0) throw new ArgumentException("Invalid duplicate key rule");
            var writeMode = Text(options, "writeMode");
            if (Array.IndexOf(new[] { "static", "formula" }, writeMode) < 0) throw new ArgumentException("Invalid lookup write mode");
            return String.Join(":", new[] { "Lookup", srcSheet, srcKey, String.Join(",", retCols), tgtKey, tgtWriteSpec, dupRule, writeMode });
        }

        static string ParseAppendSpec(Dictionary<string, object> options)
        {
            Keys(options, "sheets", "colMode", "summaryRowRule", "sourceTagMode", "targetSheetName", "colMap");
            var sheets = SheetNameList(options, "sheets", 2, 32);
            var colMode = options.ContainsKey("colMode") ? Text(options, "colMode").Trim().ToLowerInvariant() : "union";
            if (Array.IndexOf(new[] { "union", "common" }, colMode) < 0) throw new ArgumentException("Invalid colMode (expected union or common)");
            var summaryRowRule = options.ContainsKey("summaryRowRule") ? Text(options, "summaryRowRule").Trim().ToLowerInvariant() : "exclude";
            if (Array.IndexOf(new[] { "exclude", "keep" }, summaryRowRule) < 0) throw new ArgumentException("Invalid summaryRowRule (expected exclude or keep)");
            var sourceTagMode = options.ContainsKey("sourceTagMode") ? Text(options, "sourceTagMode").Trim().ToLowerInvariant() : "sheet";
            if (Array.IndexOf(new[] { "sheet", "sheet_row", "none" }, sourceTagMode) < 0) throw new ArgumentException("Invalid sourceTagMode (expected sheet, sheet_row, or none)");
            var targetSheetName = options.ContainsKey("targetSheetName") && Text(options, "targetSheetName").Trim().Length > 0
                ? ValidateSingleSheetName(Text(options, "targetSheetName"), "targetSheetName")
                : "多表追加汇总";
            if (targetSheetName.IndexOf(',') >= 0) throw new ArgumentException("Invalid worksheet name: targetSheetName");
            var colMap = options.ContainsKey("colMap") ? ValidateColMap(Text(options, "colMap")) : "";
            return String.Join(":", new[] { "Append", String.Join(",", sheets), colMode, summaryRowRule, sourceTagMode, targetSheetName, colMap });
        }

        static string ValidateHeaderLabel(string raw, string name, string defaultVal)
        {
            var s = (raw ?? "").Trim();
            if (s.Length == 0) s = defaultVal;
            if (s.Length > 64 || s.IndexOfAny(new[] { ':', '|', '\r', '\n', '\0' }) >= 0)
                throw new ArgumentException("Invalid header label: " + name);
            return s;
        }

        static string ParseUnpivotSpec(Dictionary<string, object> options)
        {
            Keys(options, "idCols", "valCols", "attrHeader", "valHeader", "emptyPolicy", "summaryRule", "targetSheetName");
            var idCols = ColLetterList(options, "idCols", 1, 16);
            var valCols = ColLetterList(options, "valCols", 1, 64);
            var idSet = new HashSet<string>(idCols, StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < valCols.Length; i++)
                if (idSet.Contains(valCols[i]))
                    throw new ArgumentException("idCols and valCols cannot overlap: " + valCols[i]);
            var attrHeader = options.ContainsKey("attrHeader") ? ValidateHeaderLabel(Text(options, "attrHeader"), "attrHeader", "属性类别") : "属性类别";
            var valHeader = options.ContainsKey("valHeader") ? ValidateHeaderLabel(Text(options, "valHeader"), "valHeader", "数值") : "数值";
            if (String.Equals(attrHeader, valHeader, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("attrHeader and valHeader must differ");
            var emptyPolicy = options.ContainsKey("emptyPolicy") ? Text(options, "emptyPolicy").Trim().ToLowerInvariant() : "skip_blank";
            if (Array.IndexOf(new[] { "skip_blank", "skip_blank_zero", "keep_all" }, emptyPolicy) < 0)
                throw new ArgumentException("Invalid emptyPolicy");
            var summaryRule = options.ContainsKey("summaryRule") ? Text(options, "summaryRule").Trim().ToLowerInvariant() : "exclude";
            if (Array.IndexOf(new[] { "exclude", "keep" }, summaryRule) < 0)
                throw new ArgumentException("Invalid summaryRule");
            var targetSheetName = options.ContainsKey("targetSheetName") && Text(options, "targetSheetName").Trim().Length > 0
                ? ValidateSingleSheetName(Text(options, "targetSheetName"), "targetSheetName")
                : "宽表转长表明细";
            if (targetSheetName.IndexOf(',') >= 0) throw new ArgumentException("Invalid worksheet name: targetSheetName");
            return String.Join(":", new[] { "Unpivot", String.Join(",", idCols), String.Join(",", valCols), attrHeader, valHeader, emptyPolicy, summaryRule, targetSheetName });
        }

        static string ParseSplitSheetSpec(Dictionary<string, object> options)
        {
            Keys(options, "splitCol", "blankKeyRule", "summaryRule", "sheetPrefix");
            var splitCol = ColLetter(options, "splitCol");
            var blankKeyRule = options.ContainsKey("blankKeyRule") ? Text(options, "blankKeyRule").Trim().ToLowerInvariant() : "bucket";
            if (Array.IndexOf(new[] { "bucket", "skip", "abort" }, blankKeyRule) < 0)
                throw new ArgumentException("Invalid blankKeyRule");
            var summaryRule = options.ContainsKey("summaryRule") ? Text(options, "summaryRule").Trim().ToLowerInvariant() : "exclude";
            if (Array.IndexOf(new[] { "exclude", "keep" }, summaryRule) < 0)
                throw new ArgumentException("Invalid summaryRule");
            var sheetPrefix = options.ContainsKey("sheetPrefix") ? Text(options, "sheetPrefix").Trim() : "";
            if (sheetPrefix.Length > 12 || sheetPrefix.IndexOfAny(new[] { ':', '|', '\r', '\n', '\0', '\\', '/', '?', '*', '[', ']', ',' }) >= 0)
                throw new ArgumentException("Invalid sheetPrefix");
            return String.Join(":", new[] { "SplitSheet", splitCol, blankKeyRule, summaryRule, sheetPrefix });
        }

        static string ParseReconcileSpec(Dictionary<string, object> options)
        {
            Keys(options, "otherSheet", "otherCol", "hasHeader");
            var otherSheet = SheetName(options, "otherSheet");
            var otherCol = ColLetter(options, "otherCol");
            var hasHeader = options.ContainsKey("hasHeader") ? Text(options, "hasHeader") : "true";
            if (hasHeader != "true" && hasHeader != "false") throw new ArgumentException("Invalid hasHeader option");
            return String.Join(":", new[] { "Reconcile", otherSheet, otherCol, hasHeader });
        }

        public static BridgeRequest Parse(string wire)
        {
            if (wire == null || wire.Length > 8192) throw new ArgumentException("Message too large");
            var json = new JavaScriptSerializer { MaxJsonLength = 8192, RecursionLimit = 12 };
            var obj = Object(json.DeserializeObject(wire));
            Keys(obj, "v", "id", "command", "actions", "options", "path", "license");
            if (!obj.ContainsKey("v") || !(obj["v"] is int) || (int)obj["v"] != 1) throw new ArgumentException("Protocol version 1 required");
            var request = new BridgeRequest { Id = Text(obj, "id"), Command = Text(obj, "command") };
            if (!Regex.IsMatch(request.Id, @"\A[A-Za-z0-9_-]{1,64}\z")) throw new ArgumentException("Invalid request ID");
            if (Array.IndexOf(new[] { "context", "candidate", "scan", "workbook", "license-status", "license-activate", "lookup_meta", "lookup_preview", "append_preview", "unpivot_preview", "split_sheet_preview", "prepare", "confirm", "execute", "invalidated", "telemetry", "open-path", "select-range" }, request.Command) < 0)
                throw new ArgumentException("Unknown command");
            if (request.Command == "license-activate" || request.Command == "license-status")
            {
                Keys(obj, "v", "id", "command", request.Command == "license-activate" ? "license" : "command");
                if (request.Command == "license-activate")
                {
                    request.License = Text(obj, "license");
                    if (request.License.Length == 0 || request.License.Length > 4096)
                        throw new ArgumentException("Invalid license length");
                }
                return request;
            }
            if (obj.ContainsKey("license")) throw new ArgumentException("License only belongs to activation");
            if (request.Command == "open-path" || request.Command == "select-range")
            {
                if (!obj.ContainsKey("path")) throw new ArgumentException("Path required");
                request.Path = Text(obj, "path");
                return request;
            }
            if (request.Command == "lookup_preview")
            {
                if (obj.ContainsKey("actions")) throw new ArgumentException("Actions only belong to prepare");
                if (!obj.ContainsKey("options")) throw new ArgumentException("Lookup options required");
                request.Actions = ParseLookupSpec(Object(obj["options"]));
                return request;
            }
            if (request.Command == "append_preview")
            {
                if (obj.ContainsKey("actions")) throw new ArgumentException("Actions only belong to prepare");
                if (!obj.ContainsKey("options")) throw new ArgumentException("Append options required");
                request.Actions = ParseAppendSpec(Object(obj["options"]));
                return request;
            }
            if (request.Command == "unpivot_preview")
            {
                if (obj.ContainsKey("actions")) throw new ArgumentException("Actions only belong to prepare");
                if (!obj.ContainsKey("options")) throw new ArgumentException("Unpivot options required");
                request.Actions = ParseUnpivotSpec(Object(obj["options"]));
                return request;
            }
            if (request.Command == "split_sheet_preview")
            {
                if (obj.ContainsKey("actions")) throw new ArgumentException("Actions only belong to prepare");
                if (!obj.ContainsKey("options")) throw new ArgumentException("SplitSheet options required");
                request.Actions = ParseSplitSheetSpec(Object(obj["options"]));
                return request;
            }
            if (obj.ContainsKey("options")) throw new ArgumentException("Top-level options only belong to lookup_preview or append_preview");
            if (request.Command != "prepare")
            {
                if (obj.ContainsKey("actions")) throw new ArgumentException("Actions only belong to prepare");
                return request;
            }
            if (!obj.ContainsKey("actions") || !(obj["actions"] is object[])) throw new ArgumentException("Action array required");
            var actions = (object[])obj["actions"];
            if (actions.Length < 1 || actions.Length > 6) throw new ArgumentException("1 to 6 actions required");
            var ids = new HashSet<string>();
            var canonical = new List<string>();
            for (int i = 0; i < actions.Length; i++)
            {
                var action = Object(actions[i]);
                Keys(action, "id", "options");
                var id = Text(action, "id");
                if (!ids.Add(id)) throw new ArgumentException("Duplicate action");
                var options = action.ContainsKey("options") ? Object(action["options"]) : new Dictionary<string, object>();
                if (id == "Transform")
                {
                    Keys(options, "mode");
                    var mode = Text(options, "mode");
                    if (Array.IndexOf(new[] { "Phone", "ID", "Name" }, mode) < 0) throw new ArgumentException("Unknown masking mode");
                    canonical.Add("Transform:" + mode);
                }
                else if (id == "Lookup")
                {
                    canonical.Add(ParseLookupSpec(options));
                }
                else if (id == "Append")
                {
                    canonical.Add(ParseAppendSpec(options));
                }
                else if (id == "Unpivot")
                {
                    canonical.Add(ParseUnpivotSpec(options));
                }
                else if (id == "SplitSheet")
                {
                    canonical.Add(ParseSplitSheetSpec(options));
                }
                else if (id == "Reconcile")
                {
                    canonical.Add(ParseReconcileSpec(options));
                }
                else if (Array.IndexOf(PlainActions, id) >= 0)
                {
                    Keys(options);
                    canonical.Add(id);
                }
                else
                {
                    throw new ArgumentException("Action disabled or unknown");
                }
            }
            request.Actions = String.Join("|", canonical);
            return request;
        }
    }
}
